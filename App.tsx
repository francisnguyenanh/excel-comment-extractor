
import React, { useState } from 'react';
import { 
  FileSpreadsheet, 
  Upload, 
  Download,
  AlertCircle,
  CheckCircle2,
  Loader2,
  Info,
  Languages,
  Cloud
} from 'lucide-react';
import { ExcelComment } from './types';
import { extractCommentsFromFile, generateResultExcel } from './services/excelService';
import { translateBatch, SUPPORTED_LANGUAGES, hasApiKey } from './services/translationService';

const App: React.FC = () => {
  // File Processing State
  const [selectedLocalFile, setSelectedLocalFile] = useState<File | null>(null);
  const [extractedComments, setExtractedComments] = useState<ExcelComment[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [apiKeyReady, setApiKeyReady] = useState(hasApiKey());
  
  // Translation State
  const [targetLang, setTargetLang] = useState<string>('');
  const [isTranslating, setIsTranslating] = useState(false);
  const [isDragging, setIsDragging] = useState(false);

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith('.xlsx')) {
        setSelectedLocalFile(file);
        setExtractedComments([]);
        setDownloadUrl(null);
      } else {
        alert('Vui lòng chỉ upload file Excel (.xlsx)');
      }
    }
  };

  const handleLocalFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setSelectedLocalFile(e.target.files[0]);
      setExtractedComments([]);
      setDownloadUrl(null);
    }
  };

  const processFile = async (file: File) => {
    setIsProcessing(true);
    setDownloadUrl(null);
    try {
      let comments = await extractCommentsFromFile(file);
      
      // Nếu có chọn ngôn ngữ đích, thực hiện dịch
      if (targetLang) {
        setIsTranslating(true);
        const originalTexts = comments.map(c => c.commentContent);
        
        // Sử dụng batch translation để tối ưu tốc độ và quota
        const translatedTexts = await translateBatch(originalTexts, targetLang);
        
        // Gán kết quả dịch vào comments
        comments = comments.map((c, index) => ({
          ...c,
          translatedContent: translatedTexts[index] || ''
        }));
        
        setIsTranslating(false);
      }
      
      setExtractedComments(comments);
      
      if (comments.length > 0) {
        const resultBlob = await generateResultExcel(comments);
        const url = URL.createObjectURL(resultBlob);
        setDownloadUrl(url);
      } else {
        alert('Không tìm thấy comment nào trong file này.');
      }
    } catch (err) {
      console.error(err);
      alert('Lỗi khi xử lý file Excel.');
    } finally {
      setIsProcessing(false);
      setIsTranslating(false);
    }
  };

  return (
    <div className="min-h-screen flex flex-col bg-gray-50">
      {/* Navigation Header */}
      <header className="bg-white border-b sticky top-0 z-10 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="bg-green-600 p-2 rounded-lg">
              <FileSpreadsheet className="text-white w-6 h-6" />
            </div>
            <h1 className="text-xl font-bold text-gray-800">Excel Comment Extractor</h1>
            {apiKeyReady ? (
               <span className="bg-green-100 text-green-800 text-xs px-2 py-0.5 rounded-full border border-green-200 flex items-center">
                 <CheckCircle2 size={12} className="mr-1" /> AI Ready
               </span>
            ) : (
               <span className="bg-orange-100 text-orange-800 text-xs px-2 py-0.5 rounded-full border border-orange-200 flex items-center" title="Chưa có API Key, sẽ dùng Google Translate miễn phí">
                 <AlertCircle size={12} className="mr-1" /> No API Key
               </span>
            )}
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-5xl w-full mx-auto px-4 py-8">
        {/* Phần hướng dẫn */}
        <div className="bg-blue-50 border-l-4 border-blue-500 p-4 mb-6 rounded-r-lg">
          <div className="flex items-start">
            <Info className="text-blue-600 mt-0.5 mr-3 flex-shrink-0" size={20} />
            <div className="flex-1">
              <h3 className="font-bold text-blue-900 mb-2">⚠️ Hướng dẫn quan trọng trước khi upload file</h3>
              <div className="text-sm text-blue-800 space-y-2">
                <p className="font-medium">Để trích xuất comments thành công, bạn cần chuyển đổi "Notes" (ghi chú) sang "Comments" trong Excel:</p>
                <ol className="list-decimal list-inside space-y-1 ml-2">
                  <li><strong>Mở file Excel</strong> có chứa Notes (dấu tam giác đỏ)</li>
                  <li>Chọn tab <strong>Review</strong> (Xem xét) trên thanh công cụ</li>
                  <li>Click vào <strong>Show All Comments</strong> để hiện tất cả ghi chú</li>
                  <li>Click chuột phải vào ô có Note → Chọn <strong>"Convert to Comment"</strong> hoặc <strong>"Chuyển sang Comment"</strong></li>
                  <li>Làm tương tự cho tất cả các ô có Notes</li>
                  <li><strong>Lưu file</strong> và upload lại</li>
                </ol>
                <p className="mt-3 bg-blue-100 p-2 rounded border border-blue-200">
                  <strong>💡 Lưu ý:</strong> Công cụ này chỉ hỗ trợ đọc <strong>Threaded Comments</strong> (Comments mới Excel 365/2019+), không hỗ trợ <strong>Notes</strong> (ghi chú kiểu cũ).
                </p>
              </div>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          
          {/* Left Column: Upload */}
          <div className="lg:col-span-1">
            <div className="bg-white rounded-xl shadow-sm border p-6">
              <h2 className="text-lg font-bold text-gray-800 mb-4 flex items-center">
                <Upload className="mr-2" size={20} />
                Tải file Excel
              </h2>
              
              <div className="space-y-4">
                <div 
                  className={`border-2 border-dashed rounded-xl p-8 text-center transition-all bg-gray-50/50 group ${isDragging ? 'border-green-500 bg-green-50 scale-105 shadow-md' : 'border-gray-200 hover:border-green-400'}`}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                >
                  <input 
                    type="file" 
                    id="file-upload" 
                    className="hidden" 
                    accept=".xlsx"
                    onChange={handleLocalFileSelect}
                  />
                  <label htmlFor="file-upload" className="cursor-pointer">
                    <div className={`bg-white shadow-sm w-12 h-12 rounded-full flex items-center justify-center mx-auto mb-4 transition-transform ${isDragging ? 'scale-125' : 'group-hover:scale-110'}`}>
                      <Upload className={`transition-colors ${isDragging ? 'text-green-600' : 'text-gray-400 group-hover:text-green-600'}`} />
                    </div>
                    <p className="text-sm font-medium text-gray-700">
                      {isDragging ? 'Thả file vào đây' : 'Kéo thả hoặc Click để chọn file'}
                    </p>
                    <p className="text-xs text-gray-400 mt-1">Hỗ trợ file Excel (.xlsx)</p>
                  </label>
                </div>
                

                {/* Language Selection */}
                <div className="bg-gray-50 p-4 rounded-xl border border-gray-200">
                  <label className="block text-sm font-semibold text-gray-700 mb-2 flex items-center">
                    <Languages size={16} className="mr-2" /> 
                    Ngôn ngữ đích (Dịch tự động)
                  </label>
                  <select
                    className="w-full p-2.5 border rounded-lg text-sm focus:ring-2 focus:ring-green-500 outline-none bg-white"
                    value={targetLang}
                    onChange={(e) => setTargetLang(e.target.value)}
                  >
                    <option value="">-- Không dịch --</option>
                    {SUPPORTED_LANGUAGES.map(lang => (
                      <option key={lang.code} value={lang.code}>{lang.name}</option>
                    ))}
                  </select>
                   {targetLang && (
                    <div className="mt-2 text-xs">
                      {apiKeyReady ? (
                         <p className="text-green-700 flex items-start">
                           <CheckCircle2 size={12} className="mr-1 mt-0.5 flex-shrink-0" />
                           Đang sử dụng Gemini AI (High Quality & Fast).
                         </p>
                      ) : (
                         <p className="text-gray-500 flex items-start">
                           <Info size={12} className="mr-1 mt-0.5 flex-shrink-0" />
                           Chưa có API Key. Chức năng dịch sẽ bị tạm tắt (chỉ hiện text gốc).
                         </p>
                      )}
                    </div>
                  )}
                </div>

                {selectedLocalFile && (
                  <div className="bg-green-50 rounded-lg p-4 border border-green-100">
                    <div className="flex items-center space-x-3 mb-3">
                      <FileSpreadsheet className="text-green-600 flex-shrink-0" size={20} />
                      <span className="text-sm font-medium text-green-800 truncate flex-1">{selectedLocalFile.name}</span>
                    </div>
                    <button 
                      onClick={() => processFile(selectedLocalFile)}
                      disabled={isProcessing}
                      className="w-full bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white font-bold py-2.5 px-4 rounded-lg shadow-sm transition-colors flex items-center justify-center"
                    >
                      {isProcessing ? (
                        <>
                          <Loader2 size={18} className="animate-spin mr-2" />
                          {isTranslating ? 'Đang dịch...' : 'Đang xử lý...'}
                        </>
                      ) : (
                        'Trích xuất Comments'
                      )}
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* Guide Card */}
            <div className="bg-indigo-900 rounded-xl shadow-lg p-6 text-white overflow-hidden relative mt-6">
               <div className="absolute top-0 right-0 p-4 opacity-10">
                  <Cloud size={80} />
               </div>
               <h3 className="font-bold text-lg mb-2 relative z-10">Hướng dẫn nhanh</h3>
               <ul className="text-sm space-y-3 opacity-90 relative z-10">
                 <li className="flex items-start">
                   <div className="bg-indigo-700 rounded-full w-5 h-5 flex items-center justify-center text-[10px] mr-2 mt-0.5 flex-shrink-0">1</div>
                   <span>Chọn file Excel (.xlsx) từ máy tính của bạn.</span>
                 </li>
                 <li className="flex items-start">
                   <div className="bg-indigo-700 rounded-full w-5 h-5 flex items-center justify-center text-[10px] mr-2 mt-0.5 flex-shrink-0">2</div>
                   <span>Chọn ngôn ngữ đích nếu muốn dịch tự động.</span>
                 </li>
                 <li className="flex items-start">
                   <div className="bg-indigo-700 rounded-full w-5 h-5 flex items-center justify-center text-[10px] mr-2 mt-0.5 flex-shrink-0">3</div>
                   <span>Hệ thống sẽ liệt kê comment và cho phép tải kết quả.</span>
                 </li>
               </ul>
            </div>
          </div>

          {/* Right Column: Results */}
          <div className="lg:col-span-2 space-y-6">
            <div className="bg-white rounded-xl shadow-sm border overflow-hidden flex flex-col h-full min-h-[500px]">
              <div className="px-6 py-4 border-b bg-gray-50 flex items-center justify-between">
                <div>
                  <h2 className="text-lg font-bold text-gray-800">Kết quả trích xuất</h2>
                  <p className="text-xs text-gray-500">
                    {extractedComments.length > 0 
                      ? `Tìm thấy ${extractedComments.length} comment` 
                      : 'Đang đợi dữ liệu...'}
                  </p>
                </div>
                {downloadUrl && (
                  <a 
                    href={downloadUrl} 
                    download="extracted_comments.xlsx"
                    className="flex items-center space-x-2 bg-green-600 hover:bg-green-700 text-white px-5 py-2.5 rounded-xl transition-all hover:scale-105 active:scale-95 shadow-lg shadow-green-200 text-sm font-bold animate-in fade-in slide-in-from-right-4 duration-300"
                  >
                    <Download size={18} />
                    <span>Tải Excel Kết Quả</span>
                  </a>
                )}
              </div>

              <div className="flex-1 overflow-auto bg-white p-0">
                {isProcessing ? (
                  <div className="flex flex-col items-center justify-center h-full space-y-4 py-20">
                    <div className="relative">
                       <Loader2 className="animate-spin text-green-600" size={48} />
                       <div className="absolute inset-0 flex items-center justify-center">
                         <FileSpreadsheet size={20} className="text-green-800" />
                       </div>
                    </div>
                    <p className="text-gray-500 font-medium">Đang trích xuất dữ liệu, vui lòng đợi...</p>
                  </div>
                ) : extractedComments.length > 0 ? (
                  <div className="animate-in fade-in duration-500">
                    <table className="min-w-full divide-y divide-gray-200">
                      <thead className="bg-gray-50 sticky top-0 z-10">
                        <tr>
                          <th className="px-4 py-3 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Sheet</th>
                          <th className="px-4 py-3 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Ô</th>
                          <th className="px-4 py-3 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Nội dung gốc</th>
                          <th className="px-4 py-3 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Nội dung Comment</th>
                          {targetLang && <th className="px-4 py-3 text-left text-xs font-bold text-gray-500 uppercase tracking-wider">Dịch ({targetLang})</th>}
                        </tr>
                      </thead>
                      <tbody className="bg-white divide-y divide-gray-200">
                        {extractedComments.map((comment, idx) => (
                          <tr key={idx} className="hover:bg-gray-50 transition-colors">
                            <td className="px-4 py-3 whitespace-nowrap text-xs font-medium text-gray-900">{comment.sheetName}</td>
                            <td className="px-4 py-3 whitespace-nowrap text-xs text-gray-500">{comment.cellAddress}</td>
                            <td className="px-4 py-3 text-xs text-gray-600">{comment.originalContent}</td>
                            <td className="px-4 py-3 text-xs text-gray-700">{comment.commentContent}</td>
                            {targetLang && <td className="px-4 py-3 text-xs text-blue-700 bg-blue-50/50">{comment.translatedContent || '...'}</td>}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-full py-20 text-gray-400">
                    <FileSpreadsheet size={64} strokeWidth={1} className="mb-4 opacity-20" />
                    <p className="text-sm italic">Chọn file để bắt đầu trích xuất comment</p>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="bg-white border-t py-6 mt-10">
         <div className="max-w-5xl mx-auto px-4 text-center">
            <p className="text-sm text-gray-400">© 2026 Excel Comment Extractor - Trích xuất Comments từ Excel nhanh chóng và dễ dàng</p>
         </div>
      </footer>
    </div>
  );
};

export default App;
