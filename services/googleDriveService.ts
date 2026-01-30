
import { DriveFile, ExcelComment } from '../types';

/**
 * Note: This service assumes an access token is available from the GIS flow.
 */
export const listExcelFiles = async (accessToken: string): Promise<DriveFile[]> => {
  const response = await fetch(
    'https://www.googleapis.com/drive/v3/files?q=mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" or mimeType="application/vnd.google-apps.spreadsheet"&fields=files(id, name, mimeType)',
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    throw new Error('Failed to fetch files from Google Drive');
  }

  const data = await response.json();
  return data.files || [];
};

export const downloadDriveFile = async (accessToken: string, fileId: string, mimeType: string): Promise<Blob> => {
  let url = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
  
  // If it's a native Google Sheet, we must export it as XLSX first
  if (mimeType === 'application/vnd.google-apps.spreadsheet') {
    url = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
  }

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    throw new Error('Failed to download file from Google Drive');
  }

  return await response.blob();
};

/**
 * Kiểm tra xem file có phải là Google Sheet native không
 */
export const isGoogleSheet = (mimeType: string): boolean => {
  return mimeType === 'application/vnd.google-apps.spreadsheet';
};

/**
 * Lấy thông tin metadata của file từ Google Drive
 */
export const getFileMetadata = async (accessToken: string, fileId: string): Promise<{ name: string; mimeType: string }> => {
  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files/${fileId}?fields=name,mimeType`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    throw new Error('Không tìm thấy file với ID này.');
  }

  return await response.json();
};

/**
 * Lấy danh sách tất cả sheets trong một Google Spreadsheet
 */
const getSpreadsheetSheets = async (accessToken: string, spreadsheetId: string): Promise<{ sheetId: number; title: string }[]> => {
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    console.error('Không thể lấy thông tin sheets:', await response.text());
    return [];
  }

  const data = await response.json();
  return (data.sheets || []).map((s: any) => ({
    sheetId: s.properties.sheetId,
    title: s.properties.title,
  }));
};

/**
 * Lấy giá trị của các ô trong một range
 */
const getSheetValues = async (
  accessToken: string, 
  spreadsheetId: string, 
  range: string
): Promise<Map<string, string>> => {
  const result = new Map<string, string>();
  
  try {
    const encodedRange = encodeURIComponent(range);
    const response = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    if (!response.ok) {
      return result;
    }

    const data = await response.json();
    const values = data.values || [];
    const sheetName = range.split('!')[0].replace(/^'|'$/g, '');
    
    values.forEach((row: any[], rowIndex: number) => {
      row.forEach((cellValue: any, colIndex: number) => {
        const colLetter = columnIndexToLetter(colIndex);
        const cellAddress = `${colLetter}${rowIndex + 1}`;
        result.set(`${sheetName}!${cellAddress}`, String(cellValue || ''));
      });
    });
  } catch (error) {
    console.error('Lỗi khi lấy giá trị sheet:', error);
  }
  
  return result;
};

/**
 * Chuyển đổi index cột (0-based) sang chữ cái Excel (A, B, ..., Z, AA, AB, ...)
 */
const columnIndexToLetter = (index: number): string => {
  let letter = '';
  let temp = index;
  
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  
  return letter;
};

/**
 * Trích xuất comments trực tiếp từ Google Sheets API
 * Đây là cách duy nhất để lấy comments từ Google Sheet vì export sang xlsx không bao gồm comments
 */
export const extractGoogleSheetComments = async (
  accessToken: string, 
  spreadsheetId: string
): Promise<ExcelComment[]> => {
  const comments: ExcelComment[] = [];
  
  try {
    console.log('[GoogleDrive] Đang trích xuất comments từ Google Sheet...');
    
    // 1. Lấy danh sách sheets
    const sheets = await getSpreadsheetSheets(accessToken, spreadsheetId);
    console.log(`[GoogleDrive] Tìm thấy ${sheets.length} sheet(s)`);
    
    // 2. Lấy tất cả comments từ spreadsheet sử dụng Drive API comments endpoint
    // Google Sheets API không có endpoint riêng cho comments, phải dùng Drive API
    const commentsResponse = await fetch(
      `https://www.googleapis.com/drive/v3/files/${spreadsheetId}/comments?fields=comments(id,content,author,anchor,quotedFileContent,replies)&pageSize=100`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    if (!commentsResponse.ok) {
      const errorText = await commentsResponse.text();
      console.error('[GoogleDrive] Lỗi khi lấy comments:', errorText);
      
      // Nếu không có quyền đọc comments, thử phương pháp khác
      if (commentsResponse.status === 403) {
        console.warn('[GoogleDrive] Không có quyền đọc comments. Cần scope: https://www.googleapis.com/auth/drive.readonly');
      }
      return comments;
    }

    const commentsData = await commentsResponse.json();
    const driveComments = commentsData.comments || [];
    
    console.log(`[GoogleDrive] Tìm thấy ${driveComments.length} comment(s) từ Drive API`);

    // 3. Lấy giá trị các ô để hiển thị nội dung gốc
    const allCellValues = new Map<string, string>();
    for (const sheet of sheets) {
      const range = `'${sheet.title}'!A1:ZZ1000`;
      const values = await getSheetValues(accessToken, spreadsheetId, range);
      values.forEach((v, k) => allCellValues.set(k, v));
    }

    // 4. Parse comments từ Drive API
    for (const comment of driveComments) {
      // anchor chứa thông tin vị trí comment
      // Format thường là: {"r":{"s":{"t":"CELL","c":"A","r":1}}}
      let sheetName = 'Unknown';
      let cellAddress = 'Unknown';
      
      try {
        if (comment.anchor) {
          const anchor = JSON.parse(comment.anchor);
          
          // Xử lý anchor format mới của Google
          if (anchor.r && anchor.r.s) {
            const selection = anchor.r.s;
            if (selection.t === 'CELL' || selection.c) {
              cellAddress = `${selection.c || 'A'}${selection.r || 1}`;
            }
          }
          
          // Tìm sheetId từ anchor và map sang tên sheet
          if (anchor.r && anchor.r.sid !== undefined) {
            const sheet = sheets.find(s => s.sheetId === anchor.r.sid);
            if (sheet) {
              sheetName = sheet.title;
            }
          } else if (sheets.length > 0) {
            // Mặc định là sheet đầu tiên nếu không xác định được
            sheetName = sheets[0].title;
          }
        }
      } catch (e) {
        console.warn('[GoogleDrive] Không parse được anchor:', comment.anchor);
      }

      const author = comment.author?.displayName || 'Không rõ';
      const content = comment.content || '';
      const quotedContent = comment.quotedFileContent?.value || '';
      
      // Lấy giá trị ô gốc
      let originalContent = allCellValues.get(`${sheetName}!${cellAddress}`) || quotedContent || '[Ô trống]';

      if (content.trim()) {
        comments.push({
          sheetName,
          cellAddress,
          originalContent: originalContent.trim() || '[Ô trống]',
          commentContent: content.trim(),
          author,
          status: 'N/A'
        });
      }

      // Xử lý replies (trả lời comment)
      if (comment.replies && Array.isArray(comment.replies)) {
        for (const reply of comment.replies) {
          const replyAuthor = reply.author?.displayName || 'Không rõ';
          const replyContent = reply.content || '';
          
          if (replyContent.trim()) {
            comments.push({
              sheetName,
              cellAddress,
              originalContent: `[Trả lời cho comment ở ${cellAddress}]`,
              commentContent: replyContent.trim(),
              author: replyAuthor,
              status: 'N/A'
            });
          }
        }
      }
    }

    console.log(`[GoogleDrive] Đã trích xuất ${comments.length} comment(s) từ Google Sheet`);
    
  } catch (error) {
    console.error('[GoogleDrive] Lỗi khi trích xuất comments:', error);
  }

  return comments;
};
