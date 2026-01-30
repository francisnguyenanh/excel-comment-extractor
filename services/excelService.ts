
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { ExcelComment } from '../types';

// Interface cho Threaded Comment
interface ThreadedComment {
  ref: string; // địa chỉ ô, ví dụ: "D16"
  personId: string;
  text: string;
  timestamp: string;
}

interface Person {
  id: string;
  displayName: string;
  userId?: string;
}

interface SheetRelation {
  sheetName: string;
  sheetId: string;
  rId: string;
}

/**
 * Chuyển đổi bất kỳ giá trị ô nào thành chuỗi văn bản thuần túy.
 * Tránh triệt để việc hiển thị [object Object].
 */
const getCellValueAsString = (cell: ExcelJS.Cell): string => {
  const value = cell.value;
  
  if (value === null || value === undefined) return '';

  // 1. Xử lý Rich Text (Văn bản có định dạng)
  if (typeof value === 'object' && 'richText' in value && Array.isArray(value.richText)) {
    return value.richText.map(rt => rt.text || '').join('');
  }

  // 2. Xử lý Công thức (Formula)
  if (typeof value === 'object' && 'formula' in value) {
    const result = (value as any).result;
    if (result !== undefined && result !== null) {
      if (typeof result === 'object') return getCellValueAsString({ value: result } as ExcelJS.Cell);
      return result.toString();
    }
    return `=${(value as any).formula}`;
  }

  // 3. Xử lý Link (Hyperlink)
  if (typeof value === 'object' && 'text' in value) {
    const textValue = (value as any).text;
    if (typeof textValue === 'object') return getCellValueAsString({ value: textValue } as ExcelJS.Cell);
    return textValue.toString();
  }

  // 4. Xử lý Ngày tháng (Date)
  if (value instanceof Date) {
    return value.toLocaleString();
  }

  // 5. Xử lý mảng (nếu có)
  if (Array.isArray(value)) {
    return value.map(v => (typeof v === 'object' ? JSON.stringify(v) : v)).join(', ');
  }

  // 6. Xử lý Object chung (Tránh [object Object])
  if (typeof value === 'object') {
    try {
      // Nếu có hàm toString tùy chỉnh không phải của Object.prototype
      if (value.toString !== Object.prototype.toString) {
        return value.toString();
      }
      return JSON.stringify(value);
    } catch {
      return '[Dữ liệu Object]';
    }
  }

  return value.toString();
};

/**
 * Trích xuất nội dung và tác giả từ đối tượng Note/Comment của ExcelJS.
 */
const getCommentContent = (note: string | ExcelJS.Comment | any): { text: string; author: string } => {
  if (!note) return { text: '', author: 'Không rõ' };

  if (typeof note === 'string') {
    return { text: note, author: 'Không rõ' };
  }

  let text = '';
  let author = 'Hệ thống';

  // ExcelJS thường lưu trong texts (mảng RichText)
  if (note.texts && Array.isArray(note.texts)) {
    text = note.texts.map((t: any) => t.text || '').join('');
    
    // Thử tìm tác giả: Thường là phần text đầu tiên kết thúc bằng dấu hai chấm
    const firstPart = note.texts[0]?.text || '';
    if (firstPart.includes(':')) {
      author = firstPart.split(':')[0].trim();
      // Nếu tác giả nằm trong text, có thể muốn loại bỏ nó khỏi nội dung chính
      // text = text.replace(firstPart, '').trim(); 
    }
  } 
  // Một số trường hợp note nằm trực tiếp trong thuộc tính note hoặc text
  else if (note.note) {
    text = typeof note.note === 'string' ? note.note : JSON.stringify(note.note);
  } else if (note.text) {
    text = typeof note.text === 'string' ? note.text : JSON.stringify(note.text);
  } else {
    // Cuối cùng nếu là object mà không biết cấu trúc, hãy stringify nó để không bị mất dữ liệu
    text = JSON.stringify(note);
  }

  return { text, author };
};

/**
 * Làm sạch nội dung comment, loại bỏ các phần không cần thiết như ID người dùng
 * Ví dụ: "======\nID#AAABu7X_-hw\nTên người dùng (2026-01-15)..." sẽ được làm sạch
 */
const cleanCommentContent = (text: string): string => {
  let cleaned = text;
  
  // Loại bỏ pattern: ======\nID#xxxxx hoặc ======ID#xxxxx
  // ID thường có dạng: ID#AAABu7X_-hw (chữ và số, có thể có dấu gạch ngang, gạch dưới)
  cleaned = cleaned.replace(/={3,}\s*ID#[A-Za-z0-9_-]+\s*/g, '');
  
  // Loại bỏ dấu ====== đứng riêng (nếu còn sót)
  cleaned = cleaned.replace(/^={3,}\s*/gm, '');
  
  // Loại bỏ ID#xxx đứng riêng một dòng
  cleaned = cleaned.replace(/^ID#[A-Za-z0-9_-]+\s*$/gm, '');
  
  // Loại bỏ các dòng trống thừa ở đầu
  cleaned = cleaned.replace(/^\s*\n+/, '');
  
  return cleaned.trim();
};

/**
 * Parse XML string và trả về Document
 */
const parseXML = (xmlString: string): Document => {
  const parser = new DOMParser();
  return parser.parseFromString(xmlString, 'application/xml');
};

/**
 * Trích xuất Threaded Comments từ file xlsx bằng cách parse trực tiếp XML
 */
const extractThreadedComments = async (arrayBuffer: ArrayBuffer): Promise<{
  comments: Map<string, { sheetName: string; cellAddress: string; text: string; author: string; timestamp: string }[]>;
  hasThreadedComments: boolean;
}> => {
  const result = new Map<string, { sheetName: string; cellAddress: string; text: string; author: string; timestamp: string }[]>();
  
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    
    // 1. Đọc danh sách persons (tác giả)
    const persons = new Map<string, Person>();
    const personsFile = zip.file('xl/persons/person.xml');
    if (personsFile) {
      const personsXml = await personsFile.async('string');
      const doc = parseXML(personsXml);
      const personElements = doc.getElementsByTagName('person');
      
      for (let i = 0; i < personElements.length; i++) {
        const el = personElements[i];
        const id = el.getAttribute('id') || '';
        const displayName = el.getAttribute('displayName') || 'Không rõ';
        const userId = el.getAttribute('userId') || '';
        persons.set(id, { id, displayName, userId });
      }
    }
    
    // 2. Đọc workbook.xml để lấy tên các sheet
    const sheetNames = new Map<string, string>(); // rId -> sheetName
    const workbookFile = zip.file('xl/workbook.xml');
    if (workbookFile) {
      const workbookXml = await workbookFile.async('string');
      const doc = parseXML(workbookXml);
      const sheetElements = doc.getElementsByTagName('sheet');
      
      for (let i = 0; i < sheetElements.length; i++) {
        const el = sheetElements[i];
        const name = el.getAttribute('name') || `Sheet${i + 1}`;
        const rId = el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || '';
        sheetNames.set(rId, name);
      }
    }
    
    // 3. Đọc workbook relationships để map rId -> sheet file
    const sheetFileMap = new Map<string, string>(); // sheet file path -> rId
    const workbookRelsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (workbookRelsFile) {
      const relsXml = await workbookRelsFile.async('string');
      const doc = parseXML(relsXml);
      const relElements = doc.getElementsByTagName('Relationship');
      
      for (let i = 0; i < relElements.length; i++) {
        const el = relElements[i];
        const rId = el.getAttribute('Id') || '';
        const target = el.getAttribute('Target') || '';
        if (target.includes('worksheets/')) {
          sheetFileMap.set(target.replace(/^\//, ''), rId);
        }
      }
    }
    
    // 4. Tìm và đọc tất cả các file threadedComments
    const threadedCommentsFiles: string[] = [];
    zip.forEach((relativePath) => {
      if (relativePath.includes('threadedComments/threadedComment')) {
        threadedCommentsFiles.push(relativePath);
      }
    });
    
    if (threadedCommentsFiles.length === 0) {
      return { comments: result, hasThreadedComments: false };
    }
    
    // 5. Đọc từng sheet's relationships để biết threadedComment nào thuộc sheet nào
    const sheetToThreadedComment = new Map<string, string>(); // threadedComment file -> sheet name
    
    for (let sheetIndex = 1; sheetIndex <= 20; sheetIndex++) {
      const sheetRelsFile = zip.file(`xl/worksheets/_rels/sheet${sheetIndex}.xml.rels`);
      if (sheetRelsFile) {
        const relsXml = await sheetRelsFile.async('string');
        const doc = parseXML(relsXml);
        const relElements = doc.getElementsByTagName('Relationship');
        
        for (let i = 0; i < relElements.length; i++) {
          const el = relElements[i];
          const target = el.getAttribute('Target') || '';
          if (target.includes('threadedComments/threadedComment')) {
            // Tìm tên sheet tương ứng
            const rId = sheetFileMap.get(`worksheets/sheet${sheetIndex}.xml`) || '';
            const sheetName = sheetNames.get(rId) || `Sheet${sheetIndex}`;
            const normalizedTarget = target.replace('../', 'xl/');
            sheetToThreadedComment.set(normalizedTarget, sheetName);
          }
        }
      }
    }
    
    // 6. Parse từng file threadedComments
    for (const tcFile of threadedCommentsFiles) {
      const file = zip.file(tcFile);
      if (!file) continue;
      
      const xmlContent = await file.async('string');
      const doc = parseXML(xmlContent);
      
      // Tìm sheetName cho file này
      let sheetName = sheetToThreadedComment.get(tcFile) || 'Unknown Sheet';
      
      // Nếu không tìm được qua rels, thử đoán từ tên file
      if (sheetName === 'Unknown Sheet') {
        const match = tcFile.match(/threadedComment(\d+)\.xml/);
        if (match) {
          const idx = parseInt(match[1]);
          // Lấy sheet name từ map nếu có
          for (const [rId, name] of sheetNames.entries()) {
            if (rId === `rId${idx}`) {
              sheetName = name;
              break;
            }
          }
          if (sheetName === 'Unknown Sheet') {
            sheetName = `Sheet${idx}`;
          }
        }
      }
      
      // Parse các threadedComment
      const commentElements = doc.getElementsByTagName('threadedComment');
      const sheetComments: { sheetName: string; cellAddress: string; text: string; author: string; timestamp: string }[] = [];
      
      for (let i = 0; i < commentElements.length; i++) {
        const el = commentElements[i];
        const ref = el.getAttribute('ref') || '';
        const personId = el.getAttribute('personId') || '';
        const dT = el.getAttribute('dT') || '';
        
        // Lấy text content
        const textEl = el.getElementsByTagName('text')[0];
        const text = textEl?.textContent || '';
        
        // Lấy tên tác giả
        const person = persons.get(personId);
        const author = person?.displayName || 'Không rõ';
        
        if (text.trim()) {
          sheetComments.push({
            sheetName,
            cellAddress: ref,
            text: text.trim(),
            author,
            timestamp: dT
          });
        }
      }
      
      if (sheetComments.length > 0) {
        result.set(sheetName, sheetComments);
      }
    }
    
    return { comments: result, hasThreadedComments: true };
  } catch (error) {
    console.error('[extractThreadedComments] Lỗi khi parse threaded comments:', error);
    return { comments: result, hasThreadedComments: false };
  }
};

const formatExcelDate = (isoString: string): string => {
  if (!isoString) return '';
  try {
    const date = new Date(isoString);
    if (isNaN(date.getTime())) return isoString;
    // Format: YYYY-MM-DD HH:mm:ss
    return date.toISOString().replace('T', ' ').substring(0, 19);
  } catch {
    return isoString;
  }
};

export const extractCommentsFromFile = async (file: File): Promise<ExcelComment[]> => {
  const arrayBuffer = await file.arrayBuffer();
  
  // Xác định định dạng file dựa trên tên file
  const fileName = file.name.toLowerCase();
  
  if (fileName.endsWith('.xls') && !fileName.endsWith('.xlsx') && !fileName.endsWith('.xlsm') && !fileName.endsWith('.xlsb')) {
    throw new Error('Định dạng .xls (Excel 97-2003) không được hỗ trợ đầy đủ. Vui lòng lưu file dưới định dạng .xlsx và thử lại.');
  }
  
  if (fileName.endsWith('.xlsb')) {
    throw new Error('Định dạng .xlsb (Excel Binary) không được hỗ trợ. Vui lòng lưu file dưới định dạng .xlsx và thử lại.');
  }

  const extractedComments: ExcelComment[] = [];
  
  // ========== BƯỚC 1: Trích xuất Threaded Comments (Comments mới Excel 365) ==========
  console.log('[ExcelService] Đang tìm Threaded Comments (Excel 365)...');
  const { comments: threadedCommentsMap, hasThreadedComments } = await extractThreadedComments(arrayBuffer);
  
  if (hasThreadedComments) {
    console.log(`[ExcelService] Tìm thấy Threaded Comments trong ${threadedCommentsMap.size} sheet(s).`);
    
    // Load workbook để lấy giá trị ô gốc
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    
    // Xử lý từng sheet có threaded comments
    for (const [sheetName, comments] of threadedCommentsMap.entries()) {
      const worksheet = workbook.getWorksheet(sheetName);
      
      for (const comment of comments) {
        let originalContent = '[Ô trống]';
        
        if (worksheet) {
          try {
            const cell = worksheet.getCell(comment.cellAddress);
            originalContent = getCellValueAsString(cell) || '[Ô trống]';
          } catch {
            originalContent = '[Không đọc được]';
          }
        }
        
        extractedComments.push({
          sheetName: comment.sheetName,
          cellAddress: comment.cellAddress,
          originalContent: originalContent.trim(),
          commentContent: cleanCommentContent(comment.text),
          author: comment.author,
          createdDate: formatExcelDate(comment.timestamp),
          status: 'N/A'
        });
      }
    }
  }
  
  // ========== BƯỚC 2: Trích xuất Notes cũ (qua ExcelJS) ==========
  console.log('[ExcelService] Đang tìm Notes (kiểu cũ)...');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  
  let cellsWithNotes = 0;
  const existingAddresses = new Set(extractedComments.map(c => `${c.sheetName}|${c.cellAddress}`));

  workbook.eachSheet((worksheet) => {
    const sheetName = worksheet.name;

    worksheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        // Kiểm tra xem ô này đã có trong danh sách threaded comments chưa
        const key = `${sheetName}|${cell.address}`;
        if (existingAddresses.has(key)) return;
        
        if (cell.note) {
          cellsWithNotes++;
          const { text, author } = getCommentContent(cell.note);
          const originalValue = getCellValueAsString(cell);

          if (text.trim().length > 0) {
            extractedComments.push({
              sheetName,
              cellAddress: cell.address,
              originalContent: originalValue.trim() || '[Ô trống]',
              commentContent: cleanCommentContent(text),
              author: author,
              createdDate: '',
              status: 'N/A'
            });
          }
        }
      });
    });
  });

  console.log(`[ExcelService] Tìm thấy ${cellsWithNotes} Notes kiểu cũ.`);
  console.log(`[ExcelService] Tổng cộng trích xuất được ${extractedComments.length} comments.`);
  
  if (extractedComments.length === 0) {
    console.warn(`[ExcelService] Không tìm thấy comment/note nào trong file.`);
  }

  return extractedComments;
};

export const generateResultExcel = async (comments: ExcelComment[]): Promise<Blob> => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Comment_Summary');

  // Cấu hình cột
  const columns = [
    { header: 'Tên Sheet', key: 'sheetName', width: 20 },
    { header: 'Địa chỉ ô', key: 'cellAddress', width: 10 },
    { header: 'Nội dung gốc của ô', key: 'originalContent', width: 30 },
    { header: 'Nội dung Comment', key: 'commentContent', width: 50 },
  ];

  // Nếu có nội dung dịch -> thêm cột
  if (comments.some(c => c.translatedContent)) {
    columns.push({ header: 'Dịch Comment', key: 'translatedContent', width: 60 });
  }

  worksheet.columns = columns;

  // Styling header
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1B5E20' }
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.height = 30;

  // Thêm dữ liệu và kẻ khung
  comments.forEach(comment => {
    const row = worksheet.addRow(comment);
    row.alignment = { vertical: 'top', wrapText: true };
    row.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};
