/**
 * Gemini for Excel - Taskpane Controller
 * Quản lý giao diện Sidebar, tương tác bảng tính và kết nối Google Gemini API
 */

let currentSelectionAddress = "Chưa chọn";
let isProcessing = false;

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        loadSettingsIntoUI();
        updateConnectionStatus();

        // Lắng nghe sự kiện người dùng di chuyển hoặc chọn vùng ô tính trong Excel
        Excel.run(async (context) => {
            context.workbook.onSelectionChanged.add(onSelectionChangedHandler);
            await context.sync();
            updateCurrentSelection();
        }).catch((err) => {
            console.log("Selection tracking init note:", err);
        });

        // Lắng nghe phím Enter trong ô nhập
        const promptArea = document.getElementById("userPrompt");
        if (promptArea) {
            promptArea.addEventListener("keydown", (e) => {
                if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                    e.preventDefault();
                    handleSendMessage();
                }
            });
        }
    }
});

// Cập nhật vùng ô tính đang được chọn trên màn hình
async function updateCurrentSelection() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("address");
            await context.sync();
            currentSelectionAddress = range.address;
            const indicator = document.getElementById("selectedRangeAddress");
            if (indicator) {
                indicator.innerText = currentSelectionAddress;
            }
        });
    } catch (e) {
        // Bỏ qua nếu bảng tính chưa sẵn sàng
    }
}

async function onSelectionChangedHandler() {
    await updateCurrentSelection();
}

// Chuyển đổi tab
function switchTab(tabName) {
    document.querySelectorAll(".tab-btn").forEach((btn) => btn.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach((tab) => tab.classList.remove("active"));

    const btn = Array.from(document.querySelectorAll(".tab-btn")).find((b) =>
        b.getAttribute("onclick").includes(tabName)
    );
    if (btn) btn.classList.add("active");

    const content = document.getElementById(`tab-${tabName}`);
    if (content) content.classList.add("active");
}

// Ẩn/Hiện mật khẩu API Key
function togglePasswordVisibility() {
    const input = document.getElementById("apiKeyInput");
    input.type = input.type === "password" ? "text" : "password";
}

// Xử lý chọn model
function handleModelChange() {
    const select = document.getElementById("modelSelect");
    const customInput = document.getElementById("customModelInput");
    if (select.value === "custom") {
        customInput.style.display = "block";
    } else {
        customInput.style.display = "none";
    }
}

// Tải cài đặt từ LocalStorage
function loadSettingsIntoUI() {
    const apiKey = localStorage.getItem("GEMINI_API_KEY") || "";
    const model = localStorage.getItem("GEMINI_MODEL") || "gemini-3.7-flash";
    const formulaModel = localStorage.getItem("GEMINI_FORMULA_MODEL") || "gemini-3.5-flash-lite";

    document.getElementById("apiKeyInput").value = apiKey;

    const modelSelect = document.getElementById("modelSelect");
    const customInput = document.getElementById("customModelInput");
    const options = Array.from(modelSelect.options).map((o) => o.value);

    if (options.includes(model)) {
        modelSelect.value = model;
        customInput.style.display = "none";
    } else if (model) {
        modelSelect.value = "custom";
        customInput.style.display = "block";
        customInput.value = model;
    }

    const formulaSelect = document.getElementById("formulaModelSelect");
    if (formulaSelect) formulaSelect.value = formulaModel;
}

// Lưu cấu hình
function saveSettings() {
    const apiKey = document.getElementById("apiKeyInput").value.trim();
    const modelSelect = document.getElementById("modelSelect").value;
    const customModel = document.getElementById("customModelInput").value.trim();
    const finalModel = modelSelect === "custom" ? customModel : modelSelect;
    const formulaModel = document.getElementById("formulaModelSelect").value;

    localStorage.setItem("GEMINI_API_KEY", apiKey);
    localStorage.setItem("GEMINI_MODEL", finalModel || "gemini-3.7-flash");
    localStorage.setItem("GEMINI_FORMULA_MODEL", formulaModel || "gemini-3.5-flash-lite");

    updateConnectionStatus();

    showNotice("settingsNotice", "notice-success", "Đã lưu cài đặt thành công! Cả Sidebar và công thức ô tính đều đã được cập nhật.");
}

// Kiểm tra kết nối tới Gemini API (Test Ping)
async function testConnection() {
    const apiKey = document.getElementById("apiKeyInput").value.trim();
    if (!apiKey) {
        showNotice("settingsNotice", "notice-danger", "Vui lòng nhập API Key trước khi kiểm tra!");
        return;
    }

    const modelSelect = document.getElementById("modelSelect").value;
    const customModel = document.getElementById("customModelInput").value.trim();
    const model = modelSelect === "custom" ? customModel : modelSelect;

    showNotice("settingsNotice", "notice-success", "Đang gửi tín hiệu kiểm tra tới máy chủ Google...");

    const startTime = performance.now();
    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts: [{ text: "Hãy trả lời duy nhất từ 'PONG' để kiểm tra kết nối." }] }]
            })
        });

        const latency = Math.round(performance.now() - startTime);
        const data = await response.json();

        if (response.ok) {
            showNotice("settingsNotice", "notice-success", `Kết nối thành công tới model ${model}! (Độ trễ: ${latency}ms)`);
            updateConnectionStatus(true);
        } else {
            showNotice("settingsNotice", "notice-danger", `Lỗi (${response.status}): ${data.error?.message || "Không thể kết nối"}`);
            updateConnectionStatus(false);
        }
    } catch (err) {
        showNotice("settingsNotice", "notice-danger", `Lỗi mạng: ${err.message}`);
        updateConnectionStatus(false);
    }
}

function updateConnectionStatus(forceValid) {
    const badge = document.getElementById("connectionBadge");
    const apiKey = localStorage.getItem("GEMINI_API_KEY");

    if (!apiKey) {
        badge.className = "badge badge-warning";
        badge.innerText = "Chưa có API Key";
    } else if (forceValid === true) {
        badge.className = "badge badge-success";
        badge.innerText = "Đã kết nối";
    } else if (forceValid === false) {
        badge.className = "badge badge-danger";
        badge.innerText = "Lỗi kết nối";
    } else {
        badge.className = "badge badge-success";
        badge.innerText = "Đã lưu API Key";
    }
}

function showNotice(elementId, className, message) {
    const el = document.getElementById(elementId);
    el.className = `notice ${className}`;
    el.innerText = message;
    el.style.display = "block";
}

// Xử lý gửi tin nhắn Chat
async function handleSendMessage() {
    if (isProcessing) return;
    const input = document.getElementById("userPrompt");
    const text = input.value.trim();
    if (!text) return;

    const apiKey = localStorage.getItem("GEMINI_API_KEY");
    if (!apiKey) {
        alert("Vui lòng vào tab 'Cài đặt' để dán Gemini API Key của bạn trước khi dùng!");
        switchTab("settings");
        return;
    }

    input.value = "";
    appendMessage("user", text);

    // Lấy dữ liệu vùng ô tính đang chọn (nếu có)
    let rangeContext = null;
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "address", "rowCount", "columnCount"]);
            await context.sync();

            // Chỉ đính kèm dữ liệu nếu vùng chọn không quá lớn và có nội dung
            if (range.rowCount * range.columnCount <= 2000) {
                rangeContext = {
                    address: range.address,
                    values: range.values
                };
            }
        });
    } catch (e) {
        console.log("Could not fetch selection context:", e);
    }

    isProcessing = true;
    const loadingId = appendLoadingMessage();

    try {
        const model = localStorage.getItem("GEMINI_MODEL") || "gemini-3.7-flash";
        let fullUserPrompt = text;
        if (rangeContext && rangeContext.values && rangeContext.values.length > 0) {
            fullUserPrompt = `[Dữ liệu vùng ô tính ${rangeContext.address}]:\n${JSON.stringify(rangeContext.values)}\n\n[Yêu cầu của người dùng]:\n${text}`;
        }

        const systemInstruction = `Bạn là trợ lý AI chuyên gia về Microsoft Excel và phân tích dữ liệu.
QUY TẮC BẮT BUỘC KHI TRẢ LỜI:
1. Khi tạo bảng tính: Hãy định dạng bảng Markdown chuẩn với các cột ngăn cách bằng dấu '|' (Ví dụ: | Cột A | Cột B | Thực lĩnh |).
2. Khi viết công thức: Hãy viết trực tiếp công thức Excel chuẩn tiếng Anh bắt đầu bằng dấu '=' trong ô tương ứng (Ví dụ: =SUM(B2:B10), =C2*D2, =IF(A2>0, "Đạt", "Trượt")).
3. Tuyệt đối không đặt công thức trong dấu ngoặc kép hoặc ký tự lạ để hệ thống có thể chèn và thực thi công thức trực tiếp trong Excel.
4. Trả lời súc tích, chuyên nghiệp.`;

        const resultText = await callGeminiAPI(apiKey, model, systemInstruction, fullUserPrompt);
        removeLoadingMessage(loadingId);
        appendMessage("assistant", resultText);

        // TỰ ĐỘNG CHÈN VÀO EXCEL NẾU ĐƯỢC BẬT
        const autoInsert = document.getElementById("autoInsertCheckbox")?.checked;
        if (autoInsert) {
            await parseAndInsertSmartly(resultText);
        }
    } catch (err) {
        removeLoadingMessage(loadingId);
        appendMessage("assistant", `❌ Đã xảy ra lỗi: ${err.message}`);
    } finally {
        isProcessing = false;
    }
}

// Các nút thao tác nhanh
function applyQuickAction(actionType) {
    const input = document.getElementById("userPrompt");
    switch (actionType) {
        case "analyze":
            input.value = "Hãy phân tích chi tiết các số liệu, chỉ số nổi bật, xu hướng và bất thường từ vùng dữ liệu này.";
            break;
        case "formula":
            input.value = "Hãy viết công thức Excel giải quyết yêu cầu sau (chỉ trả về công thức chuẩn bắt đầu bằng dấu =): ";
            break;
        case "clean":
            input.value = "Hãy kiểm tra và chỉ ra các lỗi định dạng, giá trị thiếu (null) hoặc dữ liệu trùng lặp trong bảng này.";
            break;
        case "summarize":
            input.value = "Hãy tóm tắt ngắn gọn 3 điểm quan trọng nhất từ bảng số liệu này.";
            break;
    }
    handleSendMessage();
}

// Gọi REST API của Google Gemini
async function callGeminiAPI(apiKey, model, systemInstruction, promptText) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const requestBody = {
        contents: [{ role: "user", parts: [{ text: promptText }] }]
    };

    if (systemInstruction) {
        requestBody.systemInstruction = {
            parts: [{ text: systemInstruction }]
        };
    }

    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody)
    });

    const data = await response.json();
    if (!response.ok) {
        throw new Error(data.error?.message || `HTTP ${response.status}`);
    }

    const candidate = data.candidates?.[0]?.content?.parts?.[0]?.text;
    return candidate || "(Không có phản hồi)";
}

// Bộ phân tích phản hồi từ Gemini: Tách bảng, tách công thức
function parseGeminiResponse(text) {
    const lines = text.split("\n");
    const tableRows = [];

    for (const line of lines) {
        const trimmed = line.trim();
        // Nhận diện dòng chứa bảng có dấu |
        if (trimmed.includes("|")) {
            // Bỏ qua dòng phân cách Markdown kiểu |---|---|
            if (/^\|?[\s\-:|]+\|?$/.test(trimmed)) {
                continue;
            }
            let rawCells = trimmed.split("|");
            // Nếu có dấu | ở đầu và cuối thì bỏ phần tử rỗng
            if (trimmed.startsWith("|")) rawCells.shift();
            if (trimmed.endsWith("|")) rawCells.pop();

            const cells = rawCells.map(c => c.trim().replace(/^`+|`+$/g, ""));
            if (cells.length > 1 || (cells.length === 1 && cells[0] !== "")) {
                tableRows.push(cells);
            }
        }
    }

    // Tìm công thức đơn lẻ (nếu không có bảng)
    let singleFormula = null;
    const formulaMatch = text.match(/(?:^|\s|`)(=[A-Z0-9_+\-*/^() ,:"'.]+)(?:`|\s|$)/m);
    if (formulaMatch) {
        singleFormula = formulaMatch[1].trim();
    }

    return {
        hasTable: tableRows.length >= 2,
        tableData: tableRows,
        formula: singleFormula
    };
}

// Chèn dữ liệu thông minh vào Sheet (Tự động nhận diện Bảng / Công thức / Văn bản)
async function parseAndInsertSmartly(text) {
    const parsed = parseGeminiResponse(text);

    if (parsed.hasTable) {
        await insertTableToExcel(parsed.tableData);
    } else if (parsed.formula) {
        await insertFormulaToExcel(parsed.formula);
    } else {
        await insertIntoActiveCell(text);
    }
}

// Chèn toàn bộ bảng 2 chiều vào Excel (kèm kích hoạt công thức thực sự)
async function insertTableToExcel(tableData) {
    if (!tableData || tableData.length === 0) return;

    try {
        await Excel.run(async (context) => {
            const activeCell = context.workbook.getActiveCell();
            activeCell.load(["rowIndex", "columnIndex"]);
            await context.sync();

            const numRows = tableData.length;
            const numCols = Math.max(...tableData.map(r => r.length));
            const targetRange = activeCell.getResizedRange(numRows - 1, numCols - 1);

            const matrix = [];
            for (let r = 0; r < numRows; r++) {
                const row = [];
                for (let c = 0; c < numCols; c++) {
                    let cellVal = (tableData[r][c] || "").trim();
                    // Làm sạch ký tự thừa: **in đậm**, `code`
                    cellVal = cellVal.replace(/\*\*/g, "").replace(/^`+|`+$/g, "").trim();

                    // Nếu là công thức có dấu bằng
                    if (cellVal.startsWith("=")) {
                        row.push(cellVal);
                    } else {
                        // Kiểm tra nếu là số thuần túy hoặc số có dấu phẩy ngăn cách
                        const cleanedNum = cellVal.replace(/,/g, "");
                        if (!isNaN(cleanedNum) && cleanedNum !== "") {
                            row.push(Number(cleanedNum));
                        } else {
                            row.push(cellVal);
                        }
                    }
                }
                matrix.push(row);
            }

            // Gán ma trận vào công thức của range (Excel tự động hiểu cả text, số và công thức =)
            targetRange.formulas = matrix;
            targetRange.format.autofitColumns();
            await context.sync();
        });
    } catch (err) {
        console.error("Lỗi khi chèn bảng:", err);
        // Fallback: chèn nội dung vào ô đang chọn nếu không resize được
        await insertIntoActiveCell(tableData.map(r => r.join("\t")).join("\n"));
    }
}

// Chèn công thức đơn lẻ vào ô đang chọn
async function insertFormulaToExcel(formulaStr) {
    try {
        await Excel.run(async (context) => {
            const activeCell = context.workbook.getActiveCell();
            let clean = formulaStr.trim().replace(/^`+|`+$/g, "");
            if (!clean.startsWith("=")) clean = "=" + clean;
            activeCell.formulas = [[clean]];
            await context.sync();
        });
    } catch (err) {
        alert("Lỗi chèn công thức: " + err.message);
    }
}

// Thêm tin nhắn vào khung chat với các nút thao tác thông minh
function appendMessage(role, text) {
    const container = document.getElementById("chatMessages");
    const msgDiv = document.createElement("div");
    msgDiv.className = `message ${role}`;

    const bubble = document.createElement("div");
    bubble.className = "bubble";
    bubble.innerText = text;
    msgDiv.appendChild(bubble);

    // Nếu là trợ lý AI, thêm nút hành động thông minh
    if (role === "assistant") {
        const actionsDiv = document.createElement("div");
        actionsDiv.className = "msg-actions";

        const parsed = parseGeminiResponse(text);

        // Nếu phát hiện có bảng, thêm nút chèn toàn bộ bảng
        if (parsed.hasTable) {
            const tableBtn = document.createElement("button");
            tableBtn.className = "btn-mini";
            tableBtn.style.color = "#137333";
            tableBtn.style.fontWeight = "600";
            tableBtn.innerText = `📊 Chèn bảng (${parsed.tableData.length} dòng)`;
            tableBtn.onclick = () => insertTableToExcel(parsed.tableData);
            actionsDiv.appendChild(tableBtn);
        }

        // Nếu phát hiện có công thức
        if (parsed.formula) {
            const formulaBtn = document.createElement("button");
            formulaBtn.className = "btn-mini";
            formulaBtn.style.color = "#1a73e8";
            formulaBtn.style.fontWeight = "600";
            formulaBtn.innerText = `⚡ Chèn công thức vào ô`;
            formulaBtn.onclick = () => insertFormulaToExcel(parsed.formula);
            actionsDiv.appendChild(formulaBtn);
        }

        const copyBtn = document.createElement("button");
        copyBtn.className = "btn-mini";
        copyBtn.innerText = "📋 Sao chép";
        copyBtn.onclick = () => copyText(text);

        const insertBtn = document.createElement("button");
        insertBtn.className = "btn-mini";
        insertBtn.innerText = "📥 Chèn vào ô tính";
        insertBtn.onclick = () => parseAndInsertSmartly(text);

        actionsDiv.appendChild(copyBtn);
        actionsDiv.appendChild(insertBtn);
        msgDiv.appendChild(actionsDiv);
    }

    container.appendChild(msgDiv);
    container.scrollTop = container.scrollHeight;
}

function appendLoadingMessage() {
    const container = document.getElementById("chatMessages");
    const id = "loading-" + Date.now();
    const msgDiv = document.createElement("div");
    msgDiv.id = id;
    msgDiv.className = "message assistant";
    msgDiv.innerHTML = `<div class="bubble" style="font-style:italic; color:#5f6368;">⏳ Gemini đang suy nghĩ và tính toán...</div>`;
    container.appendChild(msgDiv);
    container.scrollTop = container.scrollHeight;
    return id;
}

function removeLoadingMessage(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
}

// Chèn văn bản vào ô đang chọn trong Excel
async function insertIntoActiveCell(text) {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getActiveCell();
            if (text.trim().startsWith("=")) {
                range.formulas = [[text.trim()]];
            } else {
                range.values = [[text]];
            }
            await context.sync();
        });
    } catch (err) {
        alert("Lỗi khi ghi dữ liệu vào Excel: " + err.message);
    }
}

// Sao chép nội dung
function copyText(text) {
    navigator.clipboard.writeText(text).then(() => {
        alert("Đã sao chép vào bộ nhớ tạm!");
    }).catch(() => {
        const textarea = document.createElement("textarea");
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand("copy");
        document.body.removeChild(textarea);
        alert("Đã sao chép!");
    });
}

// Xóa cache kết quả hàm
function clearLocalCache() {
    if (typeof functionCache !== "undefined") {
        functionCache.clear();
    }
    alert("Đã làm trống bộ nhớ đệm thành công!");
}
