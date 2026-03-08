import os
import re
import uuid
import zipfile
import xml.etree.ElementTree as ET

from flask import Flask, after_this_request, render_template, request, send_file, abort
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx'}


def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# ---------------------------------------------------------------------------
# XML namespace constants
# ---------------------------------------------------------------------------
NS_WB  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_TC  = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments'


# ---------------------------------------------------------------------------
# Step 1 – Map sheet path → (sheet_name, sheet_index)
# ---------------------------------------------------------------------------
def get_sheet_name_map(z: zipfile.ZipFile) -> dict:
    """
    Returns:
        { 'xl/worksheets/sheet3.xml': ('ER_0209_入庫対象リスト', 2) }
    Index is 0-based, preserving workbook order.
    """
    # Parse xl/workbook.xml
    with z.open('xl/workbook.xml') as f:
        wb_tree = ET.parse(f)
    sheets = wb_tree.findall(f'.//{{{NS_WB}}}sheet')

    # Parse xl/_rels/workbook.xml.rels
    with z.open('xl/_rels/workbook.xml.rels') as f:
        rels_tree = ET.parse(f)
    rels = {
        r.get('Id'): r.get('Target')
        for r in rels_tree.findall(f'.//{{{NS_REL}}}Relationship')
    }

    result = {}
    for idx, sheet in enumerate(sheets):
        rid = sheet.get(f'{{{NS_R}}}id')
        target = rels.get(rid, '')          # e.g. "worksheets/sheet3.xml"
        full_path = f'xl/{target}'
        result[full_path] = (sheet.get('name'), idx)

    return result


# ---------------------------------------------------------------------------
# Step 2 – Find threadedComment XML for a given sheet
# ---------------------------------------------------------------------------
def get_threaded_comment_file(z: zipfile.ZipFile, sheet_path: str):
    """
    sheet_path: 'xl/worksheets/sheet3.xml'
    Returns the path inside the zip to the threadedComment XML, or None.
    """
    sheet_filename = sheet_path.split('/')[-1]          # 'sheet3.xml'
    rels_path = f'xl/worksheets/_rels/{sheet_filename}.rels'

    if rels_path not in z.namelist():
        return None

    with z.open(rels_path) as f:
        tree = ET.parse(f)

    for rel in tree.findall(f'.//{{{NS_REL}}}Relationship'):
        rel_type = rel.get('Type', '')
        if 'threadedComment' in rel_type:
            target = rel.get('Target', '')
            # target is usually "../threadedComments/threadedComment3.xml"
            return 'xl/' + target.replace('../', '')

    return None


# ---------------------------------------------------------------------------
# Step 3 – Parse raw comment text into individual replies
# ---------------------------------------------------------------------------
def parse_replies(raw_text: str) -> list:
    """
    Parse a threaded-comment text blob into a list of reply dicts:
        {comment_id, author, datetime, content}
    """
    # Normalize Windows CRLF → LF (XML inside .xlsx uses \r\n)
    raw_text = raw_text.replace('\r\n', '\n').replace('\r', '\n')

    # Strip Excel compatibility header that some clients prepend
    if '[Threaded comment]' in raw_text:
        idx = raw_text.find('======')
        if idx >= 0:
            raw_text = raw_text[idx:]

    # Remove leading "======" separator
    raw_text = raw_text.lstrip('=').strip()

    # Split on reply separators (4+ dashes on their own line)
    segments = re.split(r'\n[-]{4,}\n', raw_text)

    replies = []
    for segment in segments:
        lines = segment.strip().splitlines()
        if not lines:
            continue

        comment_id = ''
        author = ''
        dt_str = ''
        i = 0

        # Line 1: ID#...
        if lines[i].startswith('ID#'):
            comment_id = lines[i].strip()
            i += 1

        # Line 2: "Author Name    (YYYY-MM-DD HH:MM:SS)"
        if i < len(lines):
            m = re.match(
                r'^(.+?)\s{2,}\((\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\)\s*$',
                lines[i]
            )
            if m:
                author = m.group(1).strip()
                dt_str = m.group(2)
                i += 1

        content = '\n'.join(lines[i:]).strip()

        replies.append({
            'comment_id': comment_id,
            'author': author,
            'datetime': dt_str,
            'content': content,
        })

    return replies


# ---------------------------------------------------------------------------
# Main extraction function
# ---------------------------------------------------------------------------
def extract_comments(filepath: str) -> list:
    """
    Open an .xlsx file and return all threaded comments as a list of dicts.
    Each dict shape:
        sheet_name, cell_ref, cell_value, comment_text,
        replies, first_datetime, is_resolved, sheet_index
    """
    # Load workbook once with openpyxl for cell values
    wb = load_workbook(filepath, data_only=True)

    results = []

    with zipfile.ZipFile(filepath, 'r') as z:
        # Build sheet-path → (name, index) map
        sheet_map = get_sheet_name_map(z)

        for sheet_path, (sheet_name, sheet_index) in sheet_map.items():
            tc_path = get_threaded_comment_file(z, sheet_path)
            if tc_path is None or tc_path not in z.namelist():
                continue

            # Get the openpyxl worksheet for cell-value lookup
            ws = wb[sheet_name] if sheet_name in wb.sheetnames else None

            with z.open(tc_path) as f:
                tc_tree = ET.parse(f)

            # Try both namespace variants
            tc_elements = tc_tree.findall(f'.//{{{NS_TC}}}threadedComment')
            if not tc_elements:
                # Fallback: no-namespace or different namespace
                tc_elements = tc_tree.findall('.//threadedComment')

            for tc_el in tc_elements:
                cell_ref = tc_el.get('ref', '')
                done_attr = tc_el.get('done', '0')
                is_resolved = (done_attr == '1')

                # Raw comment text — use explicit None check, not 'or' (Element can be falsy)
                text_el = tc_el.find(f'{{{NS_TC}}}text')
                if text_el is None:
                    text_el = tc_el.find('text')
                raw_text = (text_el.text or '') if text_el is not None else ''

                replies = parse_replies(raw_text)

                # Cell value via openpyxl
                cell_value = None
                if ws is not None and cell_ref:
                    try:
                        cell_value = ws[cell_ref].value
                    except Exception:
                        cell_value = None
                if cell_value is not None:
                    cell_value = str(cell_value)

                first_datetime = replies[0]['datetime'] if replies else ''

                results.append({
                    'sheet_name': sheet_name,
                    'cell_ref': cell_ref,
                    'cell_value': cell_value,
                    'comment_text': raw_text,
                    'replies': replies,
                    'first_datetime': first_datetime,
                    'is_resolved': is_resolved,
                    'sheet_index': sheet_index,
                })

    # Sort: sheet order first, then chronological
    results.sort(key=lambda x: (x['sheet_index'], x['first_datetime'] or ''))
    return results


# ---------------------------------------------------------------------------
# Excel export builder
# ---------------------------------------------------------------------------
def build_excel(comments: list, output_path: str) -> None:
    """Generate a formatted .xlsx file from extracted comments."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Comments'

    header_font   = Font(bold=True, color='FFFFFF', size=10)
    header_fill   = PatternFill('solid', fgColor='1A56DB')
    resolved_fill = PatternFill('solid', fgColor='D1FAE5')
    open_fill     = PatternFill('solid', fgColor='DBEAFE')
    alt_fill      = PatternFill('solid', fgColor='F8FAFF')
    center        = Alignment(horizontal='center', vertical='top', wrap_text=True)
    wrap_top      = Alignment(vertical='top', wrap_text=True)
    thin_side     = Side(style='thin', color='D1D5DB')
    thin_border   = Border(left=thin_side, right=thin_side,
                           top=thin_side, bottom=thin_side)

    headers = [
        '#', 'Sheet', 'Ô', 'Nội dung ô', 'Trạng thái',
        'Reply #', 'Comment ID', 'Tên tác giả', 'Thời gian', 'Nội dung reply',
    ]
    col_widths = [5, 28, 7, 28, 12, 8, 22, 26, 20, 60]

    ws.append(headers)
    for col_idx, (hdr, width) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
        ws.column_dimensions[cell.column_letter].width = width

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = 'A2'

    data_row = 2
    for i, c in enumerate(comments, 1):
        status_label = 'Resolved' if c['is_resolved'] else 'Open'
        row_fill     = resolved_fill if c['is_resolved'] else open_fill

        for j, reply in enumerate(c['replies'], 1):
            ws.append([
                i,
                c['sheet_name'],
                c['cell_ref'],
                c['cell_value'] or '',
                status_label,
                j,
                reply['comment_id'],
                reply['author'],
                reply['datetime'],
                reply['content'],
            ])
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=data_row, column=col_idx)
                cell.border = thin_border
                cell.alignment = center if col_idx in (1, 3, 5, 6) else wrap_top
                if col_idx == 5:
                    cell.fill = row_fill
                elif data_row % 2 == 0:
                    cell.fill = alt_fill
            data_row += 1

    ws.auto_filter.ref = f'A1:{ws.cell(row=1, column=len(headers)).coordinate}'
    wb.save(output_path)


# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    error = None
    comments = []
    download_token = None

    if 'file' not in request.files:
        error = 'Không tìm thấy file trong request.'
        return render_template('index.html', error=error)

    file = request.files['file']

    if file.filename == '':
        error = 'Chưa chọn file.'
        return render_template('index.html', error=error)

    if not allowed_file(file.filename):
        error = 'Chỉ chấp nhận file .xlsx.'
        return render_template('index.html', error=error)

    token = str(uuid.uuid4())
    upload_path = os.path.join(UPLOAD_FOLDER, f'upload_{token}.xlsx')
    export_path = os.path.join(UPLOAD_FOLDER, f'comments_{token}.xlsx')

    try:
        file.save(upload_path)
        comments = extract_comments(upload_path)
        if comments:
            build_excel(comments, export_path)
            download_token = token
    except Exception as exc:
        error = f'Lỗi khi xử lý file: {exc}'
    finally:
        # Always delete the uploaded source file immediately
        if os.path.exists(upload_path):
            os.remove(upload_path)

    return render_template(
        'index.html',
        comments=comments,
        error=error,
        download_token=download_token,
    )


@app.route('/download/<token>')
def download(token):
    # Validate UUID format to prevent path traversal
    if not re.match(
        r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$',
        token,
    ):
        abort(400)

    filepath = os.path.join(UPLOAD_FOLDER, f'comments_{token}.xlsx')
    if not os.path.exists(filepath):
        abort(404)

    @after_this_request
    def remove_export(response):
        try:
            os.remove(filepath)
        except Exception:
            pass
        return response

    return send_file(
        filepath,
        as_attachment=True,
        download_name='comments_extracted.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


if __name__ == '__main__':
    app.run(debug=True)
