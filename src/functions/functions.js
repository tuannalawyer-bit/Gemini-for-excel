/**
 * Gemini for Excel - Custom Functions
 * Cung cấp các công thức tính toán AI trực tiếp trên ô tính Excel
 */

// Bộ nhớ đệm (Cache) để tránh gọi trùng lặp API khi Excel tự động tính toán lại
const functionCache = new Map();
const MAX_CACHE_SIZE = 500;

function getCacheKey(fnName, ...args) {
    return `${fnName}::` + JSON.stringify(args);
}

function getFromCache(key) {
    return functionCache.get(key);
}

function setCache(key, value) {
    if (functionCache.size >= MAX_CACHE_SIZE) {
        const firstKey = functionCache.keys().next().value;
        functionCache.delete(firstKey);
    }
    functionCache.set(key, value);
}

/**
 * Hàm chung gọi Google Gemini REST API
 */
async function executeGeminiPrompt(systemInstruction, userContent) {
    const apiKey = localStorage.getItem("GEMINI_API_KEY");
    if (!apiKey) {
        return "#CHƯA_CÓ_KEY: Hãy mở thẻ 'Gemini AI' trên thanh Ribbon để nhập API Key!";
    }

    // Ưu tiên model chuyên dụng cho công thức (mặc định gemini-3.5-flash-lite để siêu nhanh và rẻ)
    const model = localStorage.getItem("GEMINI_FORMULA_MODEL") || localStorage.getItem("GEMINI_MODEL") || "gemini-3.5-flash-lite";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const requestBody = {
        contents: [
            {
                role: "user",
                parts: [{ text: userContent }]
            }
        ]
    };

    if (systemInstruction) {
        requestBody.systemInstruction = {
            parts: [{ text: systemInstruction }]
        };
    }

    try {
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(requestBody)
        });

        const data = await response.json();

        if (!response.ok) {
            const msg = data.error?.message || "Lỗi gọi API";
            if (response.status === 429) {
                return "#QUOTA_EXCEEDED: Vượt quá giới hạn gọi API (Rate limit)!";
            }
            if (response.status === 400 || response.status === 403) {
                return `#KEY_INVALID: Khóa API không hợp lệ hoặc hết hạn!`;
            }
            return `#API_ERROR: ${msg}`;
        }

        const candidate = data.candidates?.[0]?.content?.parts?.[0]?.text;
        return candidate ? candidate.trim() : "";
    } catch (err) {
        return `#NETWORK_ERROR: ${err.message}`;
    }
}

/**
 * =GEMINI.ASK(prompt, [context])
 * Xử lý câu lệnh hoặc phân tích dữ liệu ô tính
 */
async function geminiAsk(prompt, context) {
    if (!prompt) return "";
    const cacheKey = getCacheKey("ASK", prompt, context);
    const cached = getFromCache(cacheKey);
    if (cached !== undefined) return cached;

    let content = prompt;
    if (context !== undefined && context !== null && context !== "") {
        const contextStr = typeof context === "object" ? JSON.stringify(context) : String(context);
        content += `\n\n[Dữ liệu ngữ cảnh]:\n${contextStr}`;
    }

    const systemInstruction = "Bạn là trợ lý AI thông minh trong Microsoft Excel. Hãy trả lời ngắn gọn, trực diện, chính xác, không vòng vo rườm rà. Nếu trả về số liệu hoặc danh sách, hãy giữ định dạng gọn gàng để hiển thị vừa vặn trong một ô tính.";
    const result = await executeGeminiPrompt(systemInstruction, content);
    setCache(cacheKey, result);
    return result;
}

/**
 * =GEMINI.TRANSLATE(text, targetLanguage)
 * Dịch văn bản sang ngôn ngữ đích
 */
async function geminiTranslate(text, targetLanguage) {
    if (!text) return "";
    const lang = targetLanguage || "tiếng Việt";
    const cacheKey = getCacheKey("TRANSLATE", text, lang);
    const cached = getFromCache(cacheKey);
    if (cached !== undefined) return cached;

    const systemInstruction = `Bạn là chuyên gia dịch thuật. Hãy dịch văn bản được cung cấp sang ${lang}. Chỉ trả về duy nhất nội dung đã dịch, không kèm lời giải thích, không thêm dấu ngoặc kép thừa.`;
    const result = await executeGeminiPrompt(systemInstruction, String(text));
    setCache(cacheKey, result);
    return result;
}

/**
 * =GEMINI.EXTRACT(text, whatToExtract)
 * Trích xuất email, số điện thoại, ngày tháng, họ tên...
 */
async function geminiExtract(text, whatToExtract) {
    if (!text || !whatToExtract) return "";
    const cacheKey = getCacheKey("EXTRACT", text, whatToExtract);
    const cached = getFromCache(cacheKey);
    if (cached !== undefined) return cached;

    const systemInstruction = `Bạn là công cụ bóc tách dữ liệu thông minh trong Excel. Nhiệm vụ của bạn là trích xuất '${whatToExtract}' từ đoạn văn bản được cung cấp. Chỉ trả về thông tin được trích xuất (nếu có nhiều kết quả, ngăn cách bằng dấu phẩy). Nếu không tìm thấy, chỉ trả về chuỗi rỗng "". Tuyệt đối không thêm lời giải thích.`;
    const result = await executeGeminiPrompt(systemInstruction, String(text));
    setCache(cacheKey, result);
    return result;
}

/**
 * =GEMINI.FORMULA(request, [sampleData])
 * Tự động tạo công thức Excel chuẩn từ mô tả
 */
async function geminiFormula(request, sampleData) {
    if (!request) return "";
    const cacheKey = getCacheKey("FORMULA", request, sampleData);
    const cached = getFromCache(cacheKey);
    if (cached !== undefined) return cached;

    let content = `Yêu cầu tạo công thức: ${request}`;
    if (sampleData !== undefined && sampleData !== null) {
        content += `\nCấu trúc dữ liệu mẫu: ${JSON.stringify(sampleData)}`;
    }

    const systemInstruction = "Bạn là chuyên gia công thức Excel. Hãy tạo công thức Excel chính xác nhất giải quyết yêu cầu của người dùng. Chỉ trả về DUY NHẤT công thức bắt đầu bằng dấu '=' (Ví dụ: =SUMIFS(C:C, A:A, \">10\")). Không đặt trong khối mã markdown, không thêm giải thích.";
    const result = await executeGeminiPrompt(systemInstruction, content);
    setCache(cacheKey, result);
    return result;
}

// Đăng ký các hàm với Office Runtime
if (typeof CustomFunctions !== "undefined") {
    CustomFunctions.associate("ASK", geminiAsk);
    CustomFunctions.associate("TRANSLATE", geminiTranslate);
    CustomFunctions.associate("EXTRACT", geminiExtract);
    CustomFunctions.associate("FORMULA", geminiFormula);
}
