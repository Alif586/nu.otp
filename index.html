<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Universal SMS OTP Extractor</title>
    <!-- XLSX Library for Excel files -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        :root {
            --primary: #1a73e8;
            --primary-hover: #1557b0;
            --success: #34a853;
            --bg-color: #f4f7f6;
            --card-bg: #ffffff;
            --border-color: #dadce0;
            --text-main: #202124;
            --text-sub: #5f6368;
        }

        body { font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 20px; background-color: var(--bg-color); color: var(--text-main); }
        
        .container { max-width: 1000px; margin: 40px auto; background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 10px 30px rgba(0,0,0,0.08); }
        
        h1 { color: var(--primary); margin-top: 0; margin-bottom: 10px; font-size: 28px; }
        .subtitle { color: var(--text-sub); margin-bottom: 30px; line-height: 1.5; }

        /* Upload Box */
        .upload-wrapper { position: relative; margin-bottom: 25px; }
        .upload-box { 
            border: 2px dashed var(--primary); 
            padding: 50px 20px; 
            text-align: center; 
            border-radius: 12px; 
            background: #f8fbff; 
            cursor: pointer; 
            transition: all 0.2s ease-in-out;
        }
        .upload-box:hover, .upload-box.drag-over { background: #e8f0fe; border-color: var(--primary-hover); transform: translateY(-2px); }
        .upload-box p { margin: 10px 0 5px; color: var(--text-sub); font-size: 16px; }
        .upload-box strong { color: var(--primary); }
        
        #file-input { display: none; }

        /* Controls */
        .controls { display: flex; gap: 10px; margin-bottom: 20px; align-items: center; flex-wrap: wrap; }
        .btn { 
            padding: 12px 24px; border: none; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 14px; transition: 0.2s; display: inline-flex; align-items: center; gap: 8px; 
        }
        .btn-primary { background: var(--primary); color: white; }
        .btn-primary:hover { background: var(--primary-hover); }
        .btn-success { background: var(--success); color: white; }
        .btn-success:hover { background: #2d9147; }
        .btn-outline { background: transparent; border: 1px solid var(--border-color); color: var(--text-sub); }
        .btn-outline:hover { background: #f1f3f4; color: var(--text-main); }
        .hidden { display: none !important; }

        /* Stats Area */
        .stats { 
            padding: 15px 20px; background: #e8f0fe; color: #1967d2; 
            border-radius: 8px; font-weight: 500; display: inline-flex; align-items: center; margin-bottom: 20px; border-left: 5px solid var(--primary);
        }

        /* Table */
        .table-container { overflow-x: auto; border: 1px solid var(--border-color); border-radius: 8px; max-height: 500px; overflow-y: auto; }
        table { width: 100%; border-collapse: collapse; min-width: 500px; }
        th { 
            text-align: left; background: #f8f9fa; padding: 15px; 
            border-bottom: 2px solid var(--border-color); font-size: 13px; 
            text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-sub); 
            position: sticky; top: 0; z-index: 10;
        }
        td { padding: 12px 15px; border-bottom: 1px solid var(--border-color); font-size: 14px; color: #333; }
        tr:last-child td { border-bottom: none; }
        tr:hover { background-color: #f8f9fa; }

        .otp-badge { 
            background: #fce8e6; color: #c5221f; padding: 4px 10px; 
            border-radius: 6px; font-family: 'Courier New', monospace; 
            font-weight: bold; font-size: 1.1em; letter-spacing: 1px; border: 1px solid #fad2cf; 
        }
        
        /* Loading Overlay */
        .loader-overlay {
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(255,255,255,0.8); z-index: 1000;
            display: flex; justify-content: center; align-items: center;
            flex-direction: column;
        }
        .spinner {
            border: 4px solid #f3f3f3; border-top: 4px solid var(--primary);
            border-radius: 50%; width: 40px; height: 40px;
            animation: spin 1s linear infinite;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

    </style>
</head>
<body>

<div class="container">
    <h1>Universal SMS OTP Extractor</h1>
    <p class="subtitle">Securely extract One-Time Passwords (OTP) from your exported SMS data files. Supports <strong>.xlsx, .xls, and .csv</strong> formats. Data is processed locally in your browser.</p>

    <div class="upload-wrapper">
        <div class="upload-box" id="drop-zone">
            <svg style="width:48px;height:48px;fill:var(--primary);margin-bottom:10px;" viewBox="0 0 24 24"><path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zM14 13v4h-4v-4H7l5-5 5 5h-3z"/></svg>
            <p id="file-label">Click to upload or Drag & Drop your file here</p>
            <strong>Excel (.xlsx, .xls) or CSV</strong>
            <input type="file" id="file-input" accept=".csv, .xlsx, .xls">
        </div>
    </div>

    <div id="stats-area" class="stats hidden"></div>

    <div class="controls hidden" id="action-bar">
        <button id="download-btn" class="btn btn-success" onclick="downloadTXT()">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            Download number|otp list
        </button>
        <button class="btn btn-outline" onclick="resetApp()">Clear / Upload New File</button>
    </div>

    <div class="table-container hidden" id="result-table-container">
        <table id="result-table">
            <thead>
                <tr>
                    <th width="30%">Phone Number</th>
                    <th width="70%">Extracted OTP</th>
                </tr>
            </thead>
            <tbody id="table-body"></tbody>
        </table>
    </div>
</div>

<!-- Loading Overlay -->
<div id="loader" class="loader-overlay hidden">
    <div class="spinner"></div>
    <p style="margin-top:15px; color:#555;">Processing file...</p>
</div>

<script>
    // --- Configuration & State ---
    const fileInput = document.getElementById('file-input');
    const dropZone = document.getElementById('drop-zone');
    const tableBody = document.getElementById('table-body');
    const resultTableContainer = document.getElementById('result-table-container');
    const downloadBtn = document.getElementById('download-btn');
    const statsArea = document.getElementById('stats-area');
    const actionBar = document.getElementById('action-bar');
    const fileLabel = document.getElementById('file-label');
    const loader = document.getElementById('loader');
    
    let extractedData = [];

    // --- Event Listeners ---
    
    // 1. Drag & Drop Visual Feedback
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files; // Assign dropped files to input
            handleFile(e.dataTransfer.files[0]);
        }
    });

    // 2. Click to Upload
    fileInput.addEventListener('change', function(e) {
        if (e.target.files.length) {
            handleFile(e.target.files[0]);
        }
    });

    // --- Main Logic ---

    function handleFile(file) {
        showLoader(true);
        const fileName = file.name;
        const extension = fileName.split('.').pop().toLowerCase();

        fileLabel.textContent = `Selected: ${fileName}`;

        const reader = new FileReader();

        reader.onload = function(event) {
            try {
                let rows = [];

                if (extension === 'csv') {
                    // Parse CSV manually for better control over delimiters
                    rows = parseCSV(event.target.result);
                } else if (extension === 'xlsx' || extension === 'xls') {
                    // Use SheetJS for Excel
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    rows = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                } else {
                    throw new Error("Unsupported file format.");
                }

                processRows(rows);
            } catch (error) {
                console.error(error);
                alert("Error reading file: " + error.message);
                showLoader(false);
            }
        };

        if (extension === 'csv') {
            reader.readAsText(file); // Read as text for CSV
        } else {
            reader.readAsArrayBuffer(file); // Read as buffer for Excel
        }
    }

    function parseCSV(text) {
        // Simple CSV parser handling comma or semicolon delimiters
        // Note: This simple parser assumes no commas inside quoted fields for simplicity in a single file.
        // For complex CSVs, the SheetJS library usually handles it too, but here is a native fallback logic.
        const lines = text.split(/\r\n|\n/);
        return lines.map(line => {
            // Detect delimiter
            const delimiter = line.includes(';') ? ';' : ',';
            return line.split(delimiter).map(cell => cell.trim().replace(/^"|"$/g, '')); 
        }).filter(row => row.length > 1); // Remove empty lines
    }

    function processRows(rows) {
        extractedData = [];
        tableBody.innerHTML = '';
        
        if (rows.length === 0) {
            alert("File appears to be empty.");
            showLoader(false);
            return;
        }

        // --- Step 1: Find Header Indices ---
        // We look for common names for Phone Number and SMS Content
        let numIdx = -1;
        let smsIdx = -1;

        // Check first 5 rows for headers
        for(let i = 0; i < Math.min(rows.length, 5); i++) {
            const row = rows[i].map(c => String(c || '').toLowerCase().trim());
            
            // Keywords for Number
            if (numIdx === -1) {
                if (row.some(c => c.includes('number') || c.includes('phone') || c.includes('mobile') || c.includes('tel'))) {
                    numIdx = row.findIndex(c => c.includes('number') || c.includes('phone') || c.includes('mobile') || c.includes('tel'));
                }
            }

            // Keywords for SMS Content
            if (smsIdx === -1) {
                if (row.some(c => c.includes('sms') || c.includes('message') || c.includes('content') || c.includes('text'))) {
                    smsIdx = row.findIndex(c => c.includes('sms') || c.includes('message') || c.includes('content') || c.includes('text'));
                }
            }

            // If both found, stop searching and assume this is the header row
            if (numIdx !== -1 && smsIdx !== -1) {
                // Skip this header row in the main loop by setting a start index
                // But if the user didn't have headers, we need to process data.
                // We will iterate all rows, but filter out the header row text later.
                break; 
            }
        }

        // If headers still not found, apply common defaults (e.g. Column 2 and 4)
        if (numIdx === -1) numIdx = 1; // 2nd column usually
        if (smsIdx === -1) smsIdx = 3; // 4th column usually

        // --- Step 2: Extract Data ---
        rows.forEach((row, index) => {
            // Safety check for row length
            if (!row || row.length <= Math.max(numIdx, smsIdx)) return;

            const rawNumber = row[numIdx];
            const rawSMS = row[smsIdx];

            // Skip if likely header row (heuristic)
            if (typeof rawNumber === 'string' && (rawNumber.toLowerCase().includes('number') || rawNumber.toLowerCase().includes('phone'))) return;

            const number = String(rawNumber || '').trim();
            const smsText = String(rawSMS || '').trim();

            if (number && number !== '0' && smsText) {
                const otp = extractOTP(smsText);
                if (otp) {
                    extractedData.push({ number, otp });
                    // Append to DOM efficiently using fragment if dataset is large, 
                    // but innerHTML is fine for typical OTP lists (< 10k rows)
                    const tr = `<tr><td>${number}</td><td><span class="otp-badge">${otp}</span></td></tr>`;
                    tableBody.insertAdjacentHTML('beforeend', tr);
                }
            }
        });

        // --- Step 3: Update UI ---
        showLoader(false);
        
        if (extractedData.length > 0) {
            resultTableContainer.classList.remove('hidden');
            actionBar.classList.remove('hidden');
            statsArea.classList.remove('hidden');
            statsArea.innerHTML = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="margin-right:8px;"><polyline points="20 6 9 17 4 12"></polyline></svg> Successfully extracted <strong>${extractedData.length}</strong> records containing valid OTPs.`;
        } else {
            alert("No valid OTPs found. Please check if your file contains columns with Phone Numbers and SMS Messages. \n\nDetected Column Indices: Number[${numIdx}], SMS[${smsIdx}]");
            resetApp();
        }
    }

    function extractOTP(text) {
        if (!text) return null;
        
        // 1. Clean text (remove invisible chars)
        const clean = text.replace(/[\u200B-\u200D\uFEFF\u200E\u200F]/g, '');
        
        // 2. Robust Regex Patterns
        
        // Pattern A: "Your code is 1234" (Look for code is: ...)
        let match = clean.match(/(?:code|otp|pin|verification|is)[:\s]+(\d{4,8})/i);
        if (match) return match[1];

        // Pattern B: "1234 is your code" (4-8 digits at start followed by "is")
        match = clean.match(/^(\d{4,8})\s+is/i);
        if (match) return match[1];

        // Pattern C: Isolated 4-8 digit number (generic fallback)
        // We prefer the longest number in this range if no keywords found
        const allNumbers = clean.match(/\d{4,8}/g);
        if (allNumbers) {
            // Filter out common years (like 1999, 2023) to reduce false positives
            const validCodes = allNumbers.filter(n => parseInt(n) < 2000 || parseInt(n) > 2100);
            if (validCodes.length > 0) {
                // Return the first match found that isn't a year
                return validCodes[0]; 
            }
            // If only years found, return the first one anyway (it might be 1999 as an OTP)
            return allNumbers[0];
        }

        return null;
    }

    function downloadTXT() {
        if (extractedData.length === 0) return;
        
        // Format: Number|OTP
        let content = extractedData.map(d => `${d.number}|${d.otp}`).join('\n');
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `extracted_otps_${new Date().toISOString().slice(0,10)}.txt`;
        a.click();
        URL.revokeObjectURL(url);
    }

    function resetApp() {
        extractedData = [];
        tableBody.innerHTML = '';
        fileInput.value = '';
        fileLabel.textContent = 'Click to upload or Drag & Drop your file here';
        resultTableContainer.classList.add('hidden');
        actionBar.classList.add('hidden');
        statsArea.classList.add('hidden');
    }

    function showLoader(show) {
        if (show) loader.classList.remove('hidden');
        else loader.classList.add('hidden');
    }
</script>

</body>
</html>
