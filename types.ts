
export interface ExcelComment {
  sheetName: string;
  cellAddress: string;
  originalContent: string;
  commentContent: string;
  translatedContent?: string; // Thêm trường dịch
  author: string;
  createdDate?: string;
  status: string;
}

export interface GoogleOAuthConfig {
  web: {
    client_id: string;
    project_id: string;
    auth_uri: string;
    token_uri: string;
    auth_provider_x509_cert_url: string;
    client_secret: string;
    redirect_uris: string[];
  };
}

export interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
}
