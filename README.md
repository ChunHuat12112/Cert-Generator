<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Offline Batch Certificate Generator</title>
    <!-- Include Libraries via CDN for parsing and generation -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <style>
        :root {
            --primary: #4f46e5;
            --primary-hover: #4338ca;
            --bg-color: #f3f4f6;
            --card-bg: #ffffff;
            --text-main: #1f2937;
            --border-color: #e5e7eb;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }
        body { background-color: var(--bg-color); color: var(--text-main); display: flex; height: 100vh; overflow: hidden; }
        
        /* Sidebar Control Panel */
        sidebar { width: 420px; background: var(--card-bg); border-right: 1px solid var(--border-color); display: flex; flex-direction: column; height: 100%; box-shadow: 4px 0 15px rgba(0,0,0,0.05); z-index: 10; }
        .sidebar-header { padding: 20px; border-bottom: 1px solid var(--border-color); background: #fafafa; }
        .sidebar-header h2 { font-size: 1.25rem; color: var(--primary); margin-bottom: 4px; }
        .sidebar-header p { font-size: 0.85rem; color: #6b7280; }
        
        .sidebar-content { padding: 20px; overflow-y: auto; flex: 1; display: flex; flex-direction: column; gap: 20px; }
        .section-box { background: #f9fafb; border: 1px solid var(--border-color); border-radius: 8px; padding: 15px; }
        .section-box h3 { font-size: 0.95rem; margin-bottom: 10px; color: #374151; display: flex; justify-content: space-between; align-items: center; }
        
        label { display: block; font-size: 0.8rem; font-weight: 600; margin-bottom: 5px; color: #4b5563; }
        input[type="file"], input[type="text"], select { width: 100%; padding: 8px 10px; border: 1px solid var(--border-color); border-radius: 6px; font-size: 0.85rem; margin-bottom: 12px; background: #fff; }
        
        .btn { background: var(--primary); color: white; border: none; padding: 10px 16px; border-radius: 6px; font-size: 0.9rem; font-weight: 600; cursor: pointer; width: 100%; transition: background 0.2s; text-align: center; }
        .btn:hover { background: var(--primary-hover); }
        .btn-secondary { background: #e5e7eb; color: #374151; margin-top: 6px; }
        .btn-secondary:hover { background: #d1d5db; }

        /* Main Workspace Preview */
        main { flex: 1; display: flex; flex-direction: column; height: 100%; background: #eef2f7; position: relative; overflow: hidden; }
        .toolbar { padding: 15px 25px; background: white; border-bottom: 1px solid var(--border-color); display: flex; justify-content: space-between; align-items: center; }
        .preview-container { flex: 1; display: flex; justify-content: center; align-items: center; padding: 30px; overflow: auto; position: relative; }
        
        .canvas-wrapper { position: relative; box-shadow: 0 10px 25px rgba(0,0,0,0.15); background: white; transform-origin: center center; }
        canvas { display: block; max-width: 100%; height: auto; }
        
        /* Draggable Interactive Overlay Text Elements */
        .draggable-text { position: absolute; cursor: move; border: 1px dashed transparent; user-select: none; padding: 2px 4px; }
        .draggable-text.selected { border-color: var(--primary); background: rgba(79, 70, 229, 0.1); }

        .status-bar { font-size: 0.8rem; color: #6b7280; margin-top: 8px; text-align: center; }
    </style>
</head>
<body>

    <!-- Sidebar Controls -->
    <sidebar>
        <div class="sidebar-header">
            <h2>Certificate Generator</h2>
            <p>Offline Batch Tool (Single-File Edition)</p>
        </div>
        <div class="sidebar-content">
            <!-- 1. Background Template -->
            <div class="section-box">
                <h3>1. Template Background</h3>
                <label>Upload Image (PNG / JPG)</label>
                <input type="file" id="bgUpload" accept="image/png, image/jpeg, image/webp">
                <button class="btn btn-secondary" id="resetDefaultTpl">Restore Default Template</button>
            </div>

            <!-- 2. Data List -->
            <div class="section-box">
                <h3>2. Participant List</h3>
                <label>Upload Excel (.xlsx/.xls) or CSV</label>
                <input type="file" id="listUpload" accept=".xlsx, .xls, .csv">
                <div class="status-bar" id="listStatus">No list loaded. Default sample active.</div>
            </div>

            <!-- 3. Field Configurator -->
            <div class="section-box">
                <h3>3. Active Text Element</h3>
                <label>Selected Text Field</label>
                <select id="fieldSelect">
                    <option value="name">NAMA (Participant Name)</option>
                    <option value="course">COURSE / TITLE</option>
                    <option value="date">DATE</option>
                </select>
                
                <label>Map Data Column</label>
                <select id="dataMapping"></select>

                <label>Font Size (<span id="fontSizeVal">24</span>px)</label>
                <input type="range" id="fontSizeRange" min="10" max="80" value="24" style="width:100%">

                <label>Text Color</label>
                <input type="color" id="fontColor" value="#1f2937" style="height:36px; padding:2px;">
            </div>

            <!-- 4. Generation -->
            <div>
                <button class="btn" id="generateBtn">Generate & Download ZIP</button>
            </div>
        </div>
    </sidebar>

    <!-- Visual Workspace -->
    <main>
        <div class="toolbar">
            <span><strong>Preview Mode:</strong> Click and drag texts on the canvas or use sidebar settings.</span>
            <span id="previewRecordInfo">Showing record: 1 / 1 (Sample Data)</span>
        </div>
        <div class="preview-container">
            <div class="canvas-wrapper" id="canvasWrapper">
                <canvas id="certCanvas" width="842" height="595"></canvas>
                <div id="overlayContainer"></div>
            </div>
        </div>
    </main>

    <script>
        // Core application state
        let state = {
            bgImage: null,
            listData: [{ name: "SAMPLE PARTICIPANT", course: "Advanced Frontend Architecture", date: "29 October 2026" }],
            headers: ["name", "course", "date"],
            currentRecordIndex: 0,
            selectedField: "name",
            fields: {
                name: { text: "SAMPLE PARTICIPANT", x: 50, y: 45, size: 32, color: "#111827", mapping: "name" },
                course: { text: "Advanced Frontend Architecture", x: 50, y: 60, size: 20, color: "#4b5563", mapping: "course" },
                date: { text: "29 October 2026", x: 20, y: 80, size: 14, color: "#6b7280", mapping: "date" }
            }
        };

        const canvas = document.getElementById('certCanvas');
        const ctx = canvas.getContext('2d');

        // Initialize Default Gradient Template Background
        function createDefaultTemplate() {
            const tempC = document.createElement('canvas');
            tempC.width = 1123; tempC.height = 794; // A4 Landscape ratio proxy
            const tCtx = tempC.getContext('2d');
            
            // Draw clean certificate background layout
            tCtx.fillStyle = "#ffffff";
            tCtx.fillRect(0, 0, tempC.width, tempC.height);
            
            tCtx.strokeStyle = "#4f46e5";
            tCtx.lineWidth = 12;
            tCtx.strokeRect(30, 30, tempC.width - 60, tempC.height - 60);
            
            tCtx.strokeStyle = "#c7d2fe";
            tCtx.lineWidth = 2;
            tCtx.strokeRect(45, 45, tempC.width - 90, tempC.height - 90);

            tCtx.fillStyle = "#1f2937";
            tCtx.font = "bold 40px sans-serif";
            tCtx.textAlign = "center";
            tCtx.fillText("CERTIFICATE OF COMPLETION", tempC.width / 2, 180);

            tCtx.font = "18px sans-serif";
            tCtx.fillStyle = "#6b7280";
            tCtx.fillText("This is proudly presented to", tempC.width / 2, 260);

            const img = new Image();
            img.onload = () => {
                state.bgImage = img;
                canvas.width = img.width;
                canvas.height = img.height;
                renderCanvas();
            };
            img.src = tempC.toDataURL();
        }

        function renderCanvas() {
            if (!state.bgImage) return;
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(state.bgImage, 0, 0);

            // Render each active text field onto canvas based on percentage or pixel mapping
            const currentRecord = state.listData[state.currentRecordIndex] || state.listData[0];

            for (let key in state.fields) {
                let f = state.fields[key];
                let displayText = f.text;
                
                if (f.mapping && currentRecord[f.mapping] !== undefined) {
                    displayText = currentRecord[f.mapping];
                }

                ctx.save();
                ctx.font = `bold ${f.size * (canvas.width / 800)}px sans-serif`;
                ctx.fillStyle = f.color;
                ctx.textAlign = "center";
                
                let posX = (f.x / 100) * canvas.width;
                let posY = (f.y / 100) * canvas.height;
                
                ctx.fillText(displayText, posX, posY);
                ctx.restore();
            }
        }

        // Event Handlers for UI Controls
        document.getElementById('bgUpload').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (event) => {
                const img = new Image();
                img.onload = () => {
                    state.bgImage = img;
                    canvas.width = img.width;
                    canvas.height = img.height;
                    renderCanvas();
                };
                img.src = event.target.result;
            };
            reader.readAsDataURL(file);
        });

        document.getElementById('listUpload').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (event) => {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet);
                
                if (json.length > 0) {
                    state.listData = json;
                    state.headers = Object.keys(json[0]);
                    state.currentRecordIndex = 0;
                    updateMappingDropdowns();
                    document.getElementById('listStatus').innerText = `Loaded ${json.length} records successfully.`;
                    document.getElementById('previewRecordInfo').innerText = `Showing record: 1 / ${json.length}`;
                    renderCanvas();
                }
            };
            reader.readAsArrayBuffer(file);
        });

        function updateMappingDropdowns() {
            const select = document.getElementById('dataMapping');
            select.innerHTML = '<option value="">-- Fixed Text / None --</option>';
            state.headers.forEach(h => {
                let opt = document.createElement('option');
                opt.value = h;
                opt.textContent = h;
                select.appendChild(opt);
            });
            syncSidebarToSelectedField();
        }

        const fieldSelect = document.getElementById('fieldSelect');
        const dataMapping = document.getElementById('dataMapping');
        const fontSizeRange = document.getElementById('fontSizeRange');
        const fontSizeVal = document.getElementById('fontSizeVal');
        const fontColor = document.getElementById('fontColor');

        fieldSelect.addEventListener('change', (e) => {
            state.selectedField = e.target.value;
            syncSidebarToSelectedField();
        });

        function syncSidebarToSelectedField() {
            let f = state.fields[state.selectedField];
            dataMapping.value = f.mapping || "";
            fontSizeRange.value = f.size;
            fontSizeVal.textContent = f.size;
            fontColor.value = f.color;
        }

        dataMapping.addEventListener('change', (e) => {
            state.fields[state.selectedField].mapping = e.target.value;
            renderCanvas();
        });

        fontSizeRange.addEventListener('input', (e) => {
            state.fields[state.selectedField].size = parseInt(e.target.value);
            fontSizeVal.textContent = e.target.value;
            renderCanvas();
        });

        fontColor.addEventListener('input', (e) => {
            state.fields[state.selectedField].color = e.target.value;
            renderCanvas();
        });

        document.getElementById('resetDefaultTpl').addEventListener('click', createDefaultTemplate);

        // Bulk Generation & ZIP Export logic using jsPDF and JSZip
        document.getElementById('generateBtn').addEventListener('click', async () => {
            const { jsPDF } = window.jspdf;
            const zip = new JSZip();
            const pdfFolder = zip.folder("pdf_certificates");

            document.getElementById('generateBtn').innerText = "Generating ZIP...";

            for (let i = 0; i < state.listData.length; i++) {
                state.currentRecordIndex = i;
                renderCanvas();

                // Convert Canvas to High-Quality JPEG Image URI
                const imgData = canvas.toDataURL('image/jpeg', 0.92);
                
                // Create PDF matching canvas layout orientation
                const pdfOrientation = canvas.width > canvas.height ? 'landscape' : 'portrait';
                const pdf = new jsPDF({
                    orientation: pdfOrientation,
                    unit: 'pt',
                    format: [canvas.width, canvas.height]
                });

                pdf.addImage(imgData, 'JPEG', 0, 0, canvas.width, canvas.height);
                const pdfBlobOutput = pdf.output('blob');

                // Derive safe file name from record data
                let record = state.listData[i];
                let nameKey = state.headers.find(h => /name|nama/i.test(h)) || state.headers[0];
                let safeName = (record[nameKey] || `certificate_${i+1}`).toString().replace(/[/\\?%*:|"<>]/g, '_');
                
                pdfFolder.file(`${safeName}.pdf`, pdfBlobOutput);
            }

            // Generate ZIP Archive and trigger client-side download
            const content = await zip.generateAsync({ type: "blob" });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = "certificates.zip";
            link.click();

            document.getElementById('generateBtn').innerText = "Generate & Download ZIP";
            state.currentRecordIndex = 0;
            renderCanvas();
        });

        // Initial setup execution
        createDefaultTemplate();
        updateMappingDropdowns();
    </script>
</body>
</html>
