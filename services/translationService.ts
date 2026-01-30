
/**
 * Dịch văn bản sử dụng Google Gemini API (High Quality)
 * Fallback: Google Translate API (Free endpoint via proxy)
 */

const FREE_TRANSLATE_API = 'https://translate.googleapis.com/translate_a/single';

const getApiKey = () => {
  return import.meta.env.VITE_GEMINI_API_KEY;
};

export const hasApiKey = (): boolean => {
  return !!getApiKey();
};

export const translateText = async (text: string, targetLang: string, sourceLang: string = 'auto'): Promise<string> => {
  if (!text || !text.trim()) return '';

  const apiKey = getApiKey();

  // Chỉ dịch nếu có API Key
  if (apiKey) {
    try {
      const prompt = `Translate the following text to language code "${targetLang}". Only return the translated text without quotes. Text: ${text}`;
      
      const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }]
          })
        }
      );

      const data = await response.json();
      
      if (data.error) {
        console.error('Gemini API Error:', data.error);
        throw new Error(data.error.message || 'Gemini API Error');
      }

      const translatedText = data.candidates?.[0]?.content?.parts?.[0]?.text;
      if (translatedText) {
        return translatedText.trim();
      }
    } catch (error) {
      console.warn('Gemini translation failed', error);
    }
  } else {
    console.log('No Gemini API Key found. Skipping translation.');
  }
  
  // Trả về text gốc nếu không có key hoặc lỗi
  return text;
};

/**
 * Dịch một danh sách văn bản (Batch translation) để tối ưu Rate Limit
 */
export const translateBatch = async (texts: string[], targetLang: string): Promise<string[]> => {
  if (!texts.length) return [];

  const apiKey = getApiKey();
  
  // Nếu dùng Gemini, có thể gửi batch (ghép thành JSON array)
  if (apiKey) {
    try {
      // Chia nhỏ batch nếu quá lớn (ví dụ 20 items/lượt) để tránh max token limit
      const BATCH_SIZE = 20;
      const results: string[] = [];

      for (let i = 0; i < texts.length; i += BATCH_SIZE) {
        const chunk = texts.slice(i, i + BATCH_SIZE);
        const prompt = `Translate the following array of texts to language code "${targetLang}". Return ONLY a valid JSON array of strings. Maintain the order. Sourse: ${JSON.stringify(chunk)}`;
        
        const response = await fetch(
          `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{ parts: [{ text: prompt }] }],
              generationConfig: { responseMimeType: "application/json" }
            })
          }
        );

        if (!response.ok) {
           const errText = await response.text();
           console.error(`Gemini Batch API Error (${response.status}):`, errText);
           throw new Error(`Gemini API Error: ${response.status}`);
        }

        const data = await response.json();
        const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
        
        if (content) {
          const parsed = JSON.parse(content);
          if (Array.isArray(parsed)) {
             results.push(...parsed);
          } else {
             // Fallback nếu model không trả về JSON array đúng
             results.push(...chunk); 
          }
        } else {
           results.push(...chunk);
        }
        
        // Delay nhẹ giữa các batch
        await new Promise(r => setTimeout(r, 1000));
      }
      
      // Nếu số lượng kết quả khớp số lượng đầu vào
      if (results.length === texts.length) {
        return results;
      }
    } catch (e) {
      console.warn('Batch translation failed, falling back to sequential...');
    }
  } else {
    console.warn('No API Key found. Skipping batch translation.');
    return texts; // Trả về text gốc nếu không có api key
  }

  // Fallback: Dịch từng cái (Free Google hoặc nếu Batch lỗi)
  // Chậm nhưng chắc
  const results = [];
  for (const text of texts) {
     results.push(await translateText(text, targetLang));
     // Delay nhẹ
     await new Promise(r => setTimeout(r, 50)); 
  }
  return results;
};

export const SUPPORTED_LANGUAGES = [
  { code: 'vi', name: 'Tiếng Việt' },
  { code: 'en', name: 'Tiếng Anh' },
  { code: 'ja', name: 'Tiếng Nhật' },
  { code: 'ko', name: 'Tiếng Hàn' },
  { code: 'zh', name: 'Tiếng Trung' },
];
