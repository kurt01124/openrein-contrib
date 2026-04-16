/* OpenRein PowerPoint Add-in
   PowerPoint-specific bridge: polls openrein agent server,
   executes PowerPoint commands via Office JS API. */

const SERVER = "http://localhost:19876";
let officeApp = null;
let polling = false;

// ---------------------------------------------------------------------------
// Office initialization
// ---------------------------------------------------------------------------

Office.onReady(function (info) {
    officeApp = info.host;
    const contextEl = document.getElementById("app-context");

    if (info.host === Office.HostType.PowerPoint) {
        contextEl.textContent = "PowerPoint";
    } else {
        log("error", "This add-in requires PowerPoint.");
    }

    log("info", "Add-in loaded. Connecting to server...");
    startPolling();
    setInterval(reportContext, 5000);
    reportContext();
});

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------

function log(type, text) {
    const messagesEl = document.getElementById("messages");
    const entry = document.createElement("div");
    entry.className = "log-entry " + type;

    const icons = { command: "\u2699\uFE0F", success: "\u2705", error: "\u274C", info: "\u2139\uFE0F" };
    const now = new Date();
    const time = now.getHours().toString().padStart(2, "0") + ":" +
                 now.getMinutes().toString().padStart(2, "0") + ":" +
                 now.getSeconds().toString().padStart(2, "0");

    entry.innerHTML =
        '<span class="log-icon">' + (icons[type] || "") + '</span>' +
        '<span class="log-time">' + time + '</span>' +
        '<span class="log-text">' + escapeHtml(text) + '</span>';

    messagesEl.appendChild(entry);
    messagesEl.scrollTop = messagesEl.scrollHeight;

    // Keep only last 100 entries
    while (messagesEl.children.length > 100) {
        messagesEl.removeChild(messagesEl.firstChild);
    }
}

function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}

function setStatus(connected, text) {
    const dot = document.getElementById("status-dot");
    const statusText = document.getElementById("status-text");
    if (dot) {
        dot.className = connected === "working" ? "working" : connected ? "connected" : "";
    }
    if (statusText) statusText.textContent = text;
}

// ---------------------------------------------------------------------------
// Polling for commands from Claude Gateway
// ---------------------------------------------------------------------------

function startPolling() {
    if (polling) return;
    polling = true;
    poll();
}

async function poll() {
    let wasConnected = false;

    while (polling) {
        try {
            const resp = await fetch(SERVER + "/api/poll");
            const data = await resp.json();

            if (!wasConnected) {
                wasConnected = true;
                setStatus(true, "openrein connected");
                log("success", "Server connection established");
            }

            if (data.id && data.command) {
                const action = data.command.action || "unknown";
                setStatus("working", "Executing: " + action);
                log("command", action + " " + formatParams(data.command.params));

                let result;
                try {
                    result = await executeOfficeCommand(data.command);
                    log("success", result.substring(0, 150));
                } catch (execErr) {
                    result = "Error: " + execErr.message;
                    log("error", result);
                }

                setStatus(true, "openrein connected");

                // Report result back
                await fetch(SERVER + "/api/result", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ id: data.id, result: result }),
                });
            }
        } catch (e) {
            if (wasConnected) {
                wasConnected = false;
                setStatus(false, "Disconnected — reconnecting...");
                log("error", "Server disconnected");
            }
        }
        await sleep(1000);
    }
}

function formatParams(params) {
    if (!params) return "";
    const keys = Object.keys(params);
    if (keys.length === 0) return "";
    const parts = [];
    for (const k of keys) {
        const v = params[k];
        if (typeof v === "string" && v.length > 30) {
            parts.push(k + '="' + v.substring(0, 30) + '..."');
        } else if (typeof v === "object") {
            parts.push(k + "=[...]");
        } else {
            parts.push(k + "=" + v);
        }
    }
    return "(" + parts.join(", ") + ")";
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ---------------------------------------------------------------------------
// PPTX ZIP utilities (edit_slide_xml, edit_slide_chart, edit_slide_master)
// ---------------------------------------------------------------------------

async function getPresentationBytes() {
    return new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(
            Office.FileType.Compressed,
            { sliceSize: 65536 },
            function (fileResult) {
                if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error(fileResult.error.message));
                    return;
                }
                const file = fileResult.value;
                const sliceCount = file.sliceCount;
                const slices = new Array(sliceCount);
                let received = 0;

                function onSlice(index) {
                    file.getSliceAsync(index, function (sr) {
                        if (sr.status !== Office.AsyncResultStatus.Succeeded) {
                            reject(new Error(sr.error.message));
                            return;
                        }
                        slices[index] = sr.value.data; // array of byte values
                        if (++received === sliceCount) {
                            file.closeAsync();
                            const bytes = [];
                            for (let i = 0; i < sliceCount; i++) {
                                const d = slices[i];
                                for (let j = 0; j < d.length; j++) bytes.push(d[j]);
                            }
                            resolve(new Uint8Array(bytes));
                        }
                    });
                }
                for (let i = 0; i < sliceCount; i++) onSlice(i);
            }
        );
    });
}

async function getSlideInfoFromZip(zip, slideNum) {
    const presXml = await zip.files['ppt/presentation.xml'].async('string');
    // Match all <p:sldId ...> elements in document order
    const regex = /<(?:[^:>]+:)?sldId\s[^>]+>/g;
    const matches = [...presXml.matchAll(regex)];
    if (slideNum < 1 || slideNum > matches.length) return null;
    const elem = matches[slideNum - 1][0];
    const idMatch = elem.match(/\bid="(\d+)"/);
    const rIdMatch = elem.match(/\br:id="([^"]+)"/);
    return {
        sldId: idMatch ? idMatch[1] : null,
        rId:   rIdMatch ? rIdMatch[1] : null,
    };
}

async function getSlidePathFromZip(zip, slideNum) {
    const info = await getSlideInfoFromZip(zip, slideNum);
    if (!info || !info.rId) return null;
    const relsXml = await zip.files['ppt/_rels/presentation.xml.rels'].async('string');
    const relRegex = /<Relationship\s[^>]+>/g;
    for (const m of relsXml.matchAll(relRegex)) {
        const elem = m[0];
        const idMatch = elem.match(/\bId="([^"]+)"/);
        const tgtMatch = elem.match(/\bTarget="([^"]+)"/);
        if (idMatch && idMatch[1] === info.rId && tgtMatch) {
            let tgt = tgtMatch[1].replace(/^\.\//, '');
            return tgt.startsWith('/') ? tgt.substring(1) : 'ppt/' + tgt;
        }
    }
    return null;
}

function getRelativeLuminance(hex) {
    const h = (hex || '').replace('#', '');
    let r, g, b;
    if (h.length === 3) {
        r = parseInt(h[0]+h[0], 16);
        g = parseInt(h[1]+h[1], 16);
        b = parseInt(h[2]+h[2], 16);
    } else if (h.length >= 6) {
        r = parseInt(h.substring(0,2), 16);
        g = parseInt(h.substring(2,4), 16);
        b = parseInt(h.substring(4,6), 16);
    } else { return 0; }
    const lin = c => {
        const s = c / 255;
        return s <= 0.04045 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
    };
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

function getContrastRatio(hex1, hex2) {
    try {
        const l1 = getRelativeLuminance(hex1);
        const l2 = getRelativeLuminance(hex2);
        const lighter = Math.max(l1, l2);
        const darker  = Math.min(l1, l2);
        return (lighter + 0.05) / (darker + 0.05);
    } catch(e) { return 21; }
}

// Replace slide N with modified version from zip
async function _editSlideZip(slideNum, code) {
    let pptxBytes;
    try { pptxBytes = await getPresentationBytes(); }
    catch(e) { return "Error reading presentation: " + e.message; }

    let zip;
    try { zip = await JSZip.loadAsync(pptxBytes); }
    catch(e) { return "Error parsing PPTX: " + e.message; }

    if (!code || !code.trim()) return "Error: code parameter is empty. Provide JavaScript code to modify the slide XML.";

    const slidePath = await getSlidePathFromZip(zip, slideNum);

    let dirty = false;
    const markDirty = () => { dirty = true; };

    try {
        const AsyncFunc = Object.getPrototypeOf(async function(){}).constructor;
        // Append markDirty() only if not already present
        const wrappedCode = code.includes('markDirty') ? code : code + '\nmarkDirty();';
        const fn = new AsyncFunc('zip', 'markDirty', 'slidePath', 'slideNum', wrappedCode);
        await fn(zip, markDirty, slidePath, slideNum);
    } catch(e) { return "Error in slide code: " + e.message; }

    if (!dirty) return "ERROR: markDirty() was not called — no changes applied. You MUST call markDirty() as the last line after zip.file(slidePath, xmlString).";

    // Validate XML before reimport
    if (slidePath) {
        try {
            const modXml = await zip.files[slidePath].async('string');
            const doc = new DOMParser().parseFromString(modXml, 'application/xml');
            const err = doc.querySelector('parsererror');
            if (err) {
                return "Error: Invalid XML — " + err.textContent.substring(0, 300)
                     + "\n\nFirst 500 chars:\n" + modXml.substring(0, 500);
            }
        } catch(e) { /* continue */ }
    }

    const info = await getSlideInfoFromZip(zip, slideNum);
    if (!info || !info.sldId) return "Error: slide " + slideNum + " not found in ZIP.";

    let newBase64;
    try { newBase64 = await zip.generateAsync({ type: 'base64' }); }
    catch(e) { return "Error generating PPTX: " + e.message; }

    return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        if (slideNum > slides.items.length)
            return "Error: presentation has only " + slides.items.length + " slides.";

        for (const s of slides.items) s.load("id");
        await context.sync();

        const origId = slides.items[slideNum - 1].id;
        const prevId = slideNum > 1 ? slides.items[slideNum - 2].id : undefined;

        const opts = { formatting: "KeepSourceFormatting", sourceSlideIds: [info.sldId] };
        if (prevId !== undefined) opts.targetSlideId = prevId;

        context.presentation.insertSlidesFromBase64(newBase64, opts);
        await context.sync();

        // Delete original by ID
        const slides2 = context.presentation.slides;
        slides2.load("items");
        await context.sync();
        for (const s of slides2.items) s.load("id");
        await context.sync();
        for (const s of slides2.items) {
            if (s.id === origId) { s.delete(); break; }
        }
        await context.sync();

        return "Slide " + slideNum + " XML updated.";
    });
}

// Replace entire presentation (for master edits)
async function _editMasterZip(code) {
    let pptxBytes;
    try { pptxBytes = await getPresentationBytes(); }
    catch(e) { return "Error reading presentation: " + e.message; }

    let zip;
    try { zip = await JSZip.loadAsync(pptxBytes); }
    catch(e) { return "Error parsing PPTX: " + e.message; }

    let dirty = false;
    const markDirty = () => { dirty = true; };

    try {
        const AsyncFunc = Object.getPrototypeOf(async function(){}).constructor;
        const wrappedCode = code.includes('markDirty') ? code : code + '\nmarkDirty();';
        const fn = new AsyncFunc('zip', 'markDirty', wrappedCode);
        await fn(zip, markDirty);
    } catch(e) { return "Error in master code: " + e.message; }

    if (!dirty) return "ERROR: markDirty() was not called — no changes applied. You MUST call markDirty() as the last line after zip.file(slidePath, xmlString).";

    // Collect all source slide IDs
    const presXml = await zip.files['ppt/presentation.xml'].async('string');
    const allSldIds = [...presXml.matchAll(/<(?:[^:>]+:)?sldId\s[^>]+>/g)]
        .map(m => { const mm = m[0].match(/\bid="(\d+)"/); return mm ? mm[1] : null; })
        .filter(Boolean);

    if (allSldIds.length === 0) return "Error: No slides found in modified ZIP.";

    let newBase64;
    try { newBase64 = await zip.generateAsync({ type: 'base64' }); }
    catch(e) { return "Error generating PPTX: " + e.message; }

    return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        for (const s of slides.items) s.load("id");
        await context.sync();

        const origIds = slides.items.map(s => s.id);
        const lastOrigId = origIds[origIds.length - 1];

        context.presentation.insertSlidesFromBase64(newBase64, {
            formatting: "KeepSourceFormatting",
            sourceSlideIds: allSldIds,
            targetSlideId: lastOrigId,
        });
        await context.sync();

        const slides2 = context.presentation.slides;
        slides2.load("items");
        await context.sync();
        for (const s of slides2.items) s.load("id");
        await context.sync();
        const toDelete = slides2.items.filter(s => origIds.includes(s.id));
        for (const s of toDelete) s.delete();
        await context.sync();

        return "Slide master updated. " + allSldIds.length + " slide(s) reloaded.";
    });
}

// Duplicate slide N immediately after itself
async function _duplicateSlideZip(slideNum) {
    let pptxBytes;
    try { pptxBytes = await getPresentationBytes(); }
    catch(e) { return "Error reading presentation: " + e.message; }

    let zip;
    try { zip = await JSZip.loadAsync(pptxBytes); }
    catch(e) { return "Error parsing PPTX: " + e.message; }

    const info = await getSlideInfoFromZip(zip, slideNum);
    if (!info || !info.sldId) return "Error: slide " + slideNum + " not found.";

    let base64;
    try { base64 = await zip.generateAsync({ type: 'base64' }); }
    catch(e) { return "Error generating PPTX: " + e.message; }

    return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        if (slideNum > slides.items.length)
            return "Error: slide " + slideNum + " doesn't exist.";

        for (const s of slides.items) s.load("id");
        await context.sync();

        const afterId = slides.items[slideNum - 1].id;
        context.presentation.insertSlidesFromBase64(base64, {
            formatting: "KeepSourceFormatting",
            sourceSlideIds: [info.sldId],
            targetSlideId: afterId,
        });
        await context.sync();
        return "Slide " + slideNum + " duplicated.";
    });
}

// ---------------------------------------------------------------------------
// Report document context to Claude Gateway
// ---------------------------------------------------------------------------

async function reportContext() {
    try {
        const context = await getDocumentContext();
        await fetch(SERVER + "/api/report_context", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(context),
        });
    } catch (e) { /* ignore */ }
}

async function getDocumentContext() {
    const ctx = {
        app: officeApp ? officeApp.toString() : "unknown",
        timestamp: Date.now(),
    };
    try {
        if (officeApp === Office.HostType.PowerPoint) {
            await PowerPoint.run(async (context) => {
                const slides = context.presentation.slides;
                slides.load("items");
                await context.sync();
                ctx.slideCount = slides.items.length;
            });
        }
    } catch (e) {
        ctx.error = e.message;
    }
    return ctx;
}

// ---------------------------------------------------------------------------
// Office JS document manipulation
// ---------------------------------------------------------------------------

async function executeOfficeCommand(command) {
    try {
        if (officeApp === Office.HostType.PowerPoint) {
            return await executePowerPointCommand(command);
        }
    } catch (e) {
        return "Error: " + e.message;
    }
    return "This add-in is PowerPoint only.";
}

async function executePowerPointCommand(cmd) {
    const p = cmd.params || {};

    // --- office_command: delegate to nested action ---
    if (cmd.action === "office_command") {
        return await executePowerPointCommand({
            action: p.action || "",
            params: p.params || {},
        });
    }

    // --- OOXML ZIP operations (use getFileAsync outside PowerPoint.run) ---
    if (cmd.action === "edit_slide_xml" || cmd.action === "edit_slide_chart") {
        return await _editSlideZip(p.slide || 1, p.code || '');
    }
    if (cmd.action === "edit_slide_master") {
        return await _editMasterZip(p.code || '');
    }
    if (cmd.action === "duplicate_slide") {
        return await _duplicateSlideZip(p.slide || 1);
    }

    return PowerPoint.run(async (context) => {
        const presentation = context.presentation;
        const p = cmd.params || {};

        async function getSlide(slideNum) {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            const idx = (slideNum || slides.items.length) - 1;
            return slides.items[idx];
        }

        if (cmd.action === "get_office_context") {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            return JSON.stringify({
                app: "PowerPoint",
                slideCount: slides.items.length,
                timestamp: Date.now(),
            });
        }

        if (cmd.action === "add_slide") {
            presentation.slides.add();
            await context.sync();
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            return "Slide added. Total: " + slides.items.length;
        }

        if (cmd.action === "add_text_box") {
            const slide = await getSlide(p.slide);
            const shape = slide.shapes.addTextBox(p.text || "", {
                left: p.left || 50,
                top: p.top || 50,
                width: p.width || 400,
                height: p.height || 50,
            });
            if (p.font_size || p.font_color || p.font_bold || p.font_name) {
                shape.textFrame.textRange.font.size = p.font_size || 18;
                if (p.font_color) shape.textFrame.textRange.font.color = p.font_color;
                if (p.font_bold) shape.textFrame.textRange.font.bold = true;
                if (p.font_name) shape.textFrame.textRange.font.name = p.font_name;
            }
            await context.sync();
            return "Text box added.";
        }

        if (cmd.action === "add_shape") {
            const slide = await getSlide(p.slide);
            const shapeType = p.shape_type || "Rectangle";
            const shape = slide.shapes.addGeometricShape(shapeType, {
                left: p.left || 50,
                top: p.top || 50,
                width: p.width || 200,
                height: p.height || 100,
            });
            if (p.fill) shape.fill.setSolidColor(p.fill);
            if (p.text) {
                shape.textFrame.textRange.text = p.text;
                if (p.font_size) shape.textFrame.textRange.font.size = p.font_size;
                if (p.font_color) shape.textFrame.textRange.font.color = p.font_color;
                if (p.font_bold) shape.textFrame.textRange.font.bold = true;
                if (p.font_name) shape.textFrame.textRange.font.name = p.font_name;
            }
            if (p.line_color) shape.lineFormat.color = p.line_color;
            if (p.no_line) shape.lineFormat.visible = false;
            await context.sync();
            return "Shape added.";
        }

        if (cmd.action === "add_image") {
            const slide = await getSlide(p.slide);
            if (p.base64) {
                slide.shapes.addImage(p.base64, {
                    left: p.left || 50,
                    top: p.top || 50,
                    width: p.width || 300,
                    height: p.height || 200,
                });
            }
            await context.sync();
            return "Image added.";
        }

        if (cmd.action === "add_table") {
            const slide = await getSlide(p.slide);
            const values = p.values || [[""]];
            const rows = values.length;
            const cols = values[0] ? values[0].length : 1;
            try {
                // Method 1: addTable with data array
                const table = slide.shapes.addTable(rows, cols, {
                    left: p.left || 50,
                    top: p.top || 50,
                    width: p.width || 600,
                    height: p.height || (rows * 40),
                    values: values,
                });
                await context.sync();
                return "Table added (" + rows + "x" + cols + ").";
            } catch(e1) {
                try {
                    // Method 2: addTable then set cell text individually
                    const table = slide.shapes.addTable(rows, cols, {
                        left: p.left || 50,
                        top: p.top || 50,
                        width: p.width || 600,
                        height: p.height || (rows * 40),
                    });
                    await context.sync();
                    // Try to set values
                    for (let r = 0; r < rows; r++) {
                        for (let c = 0; c < cols; c++) {
                            if (values[r] && values[r][c] !== undefined) {
                                try {
                                    const cell = table.rows.getItemAt(r).cells.getItemAt(c);
                                    cell.body.clear();
                                    cell.body.insertParagraph(String(values[r][c]), "Start");
                                    await context.sync();
                                } catch(ce) { /* skip cell errors */ }
                            }
                        }
                    }
                    return "Table added (" + rows + "x" + cols + ").";
                } catch(e2) {
                    // Method 3: Fallback — create text-based table using shapes
                    const cellW = (p.width || 600) / cols;
                    const cellH = 35;
                    const startLeft = p.left || 50;
                    const startTop = p.top || 50;
                    for (let r = 0; r < rows; r++) {
                        for (let c = 0; c < cols; c++) {
                            const val = (values[r] && values[r][c] !== undefined) ? String(values[r][c]) : "";
                            const cellShape = slide.shapes.addGeometricShape("Rectangle", {
                                left: startLeft + c * cellW,
                                top: startTop + r * cellH,
                                width: cellW,
                                height: cellH,
                            });
                            cellShape.fill.setSolidColor(r === 0 ? (p.header_color || "#2B579A") : "#FFFFFF");
                            cellShape.lineFormat.color = "#CCCCCC";
                            cellShape.lineFormat.weight = 0.5;
                            cellShape.textFrame.textRange.text = val;
                            cellShape.textFrame.textRange.font.size = p.font_size || 11;
                            cellShape.textFrame.textRange.font.color = r === 0 ? "#FFFFFF" : "#333333";
                            if (r === 0) cellShape.textFrame.textRange.font.bold = true;
                        }
                    }
                    await context.sync();
                    return "Table added as shapes (" + rows + "x" + cols + ").";
                }
            }
        }

        if (cmd.action === "set_background") {
            const slide = await getSlide(p.slide);
            try {
                // Try direct fill
                slide.fill.setSolidColor(p.color || "#FFFFFF");
                await context.sync();
                return "Background set.";
            } catch(e1) {
                try {
                    // Fallback: add a full-size rectangle behind everything
                    const bg = slide.shapes.addGeometricShape("Rectangle", {
                        left: 0, top: 0, width: 960, height: 540,
                    });
                    bg.fill.setSolidColor(p.color || "#FFFFFF");
                    bg.lineFormat.visible = false;
                    // Move to back — not directly supported, but it's added first
                    await context.sync();
                    return "Background set (via shape).";
                } catch(e2) {
                    return "Error setting background: " + e1.message;
                }
            }
        }

        if (cmd.action === "delete_slide") {
            const slide = await getSlide(p.slide);
            slide.delete();
            await context.sync();
            return "Slide deleted.";
        }

        if (cmd.action === "get_slide_count") {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            return "Slide count: " + slides.items.length;
        }

        if (cmd.action === "get_slide_text") {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            const slideIndex = (p.slide || 1) - 1;
            const slide = slides.items[slideIndex];
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            let text = "";
            for (const shape of shapes.items) {
                if (shape.textFrame) {
                    shape.textFrame.load("textRange");
                    await context.sync();
                    if (shape.textFrame.textRange) {
                        shape.textFrame.textRange.load("text");
                        await context.sync();
                        text += shape.textFrame.textRange.text + "\n";
                    }
                }
            }
            return text || "(no text on slide " + (slideIndex + 1) + ")";
        }

        if (cmd.action === "get_slide_image") {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            const slideIndex = (p.slide || 1) - 1;
            const slide = slides.items[slideIndex];
            const imgResult = slide.getImageAsBase64({
                width: p.width || 960,
                height: p.height || 540,
            });
            await context.sync();
            return "__IMAGE__" + imgResult.value;
        }

        // --- Shape styling ---

        if (cmd.action === "set_shape_fill") {
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            const shape = shapes.items[(p.shape_index || 1) - 1];
            shape.fill.setSolidColor(p.color || "#FFFFFF");
            await context.sync();
            return "Shape fill color set.";
        }

        if (cmd.action === "set_shape_line") {
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            const shape = shapes.items[(p.shape_index || 1) - 1];
            if (p.color) shape.lineFormat.color = p.color;
            if (p.width) shape.lineFormat.weight = p.width;
            if (p.visible === false) shape.lineFormat.visible = false;
            await context.sync();
            return "Shape line style set.";
        }

        // --- Text styling ---

        if (cmd.action === "set_text_style") {
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            const shape = shapes.items[(p.shape_index || 1) - 1];
            const tr = shape.textFrame.textRange;
            if (p.font_name) tr.font.name = p.font_name;
            if (p.font_size) tr.font.size = p.font_size;
            if (p.font_color) tr.font.color = p.font_color;
            if (p.bold !== undefined) tr.font.bold = p.bold;
            if (p.italic !== undefined) tr.font.italic = p.italic;
            if (p.underline !== undefined) tr.font.underline = p.underline ? "Single" : "None";
            await context.sync();
            return "Text style applied.";
        }

        if (cmd.action === "set_text_align") {
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            const shape = shapes.items[(p.shape_index || 1) - 1];
            const paragraphs = shape.textFrame.textRange.paragraphs;
            paragraphs.load("items");
            await context.sync();
            const alignMap = { left: "Left", center: "Center", right: "Right", justify: "Justify" };
            const align = alignMap[(p.align || "left").toLowerCase()] || "Left";
            for (const para of paragraphs.items) {
                para.horizontalAlignment = align;
            }
            await context.sync();
            return "Text alignment set to " + align + ".";
        }

        if (cmd.action === "set_line_spacing") {
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            const shape = shapes.items[(p.shape_index || 1) - 1];
            const paragraphs = shape.textFrame.textRange.paragraphs;
            paragraphs.load("items");
            await context.sync();
            for (const para of paragraphs.items) {
                para.lineSpacing = p.spacing || 1.5;
            }
            await context.sync();
            return "Line spacing set.";
        }

        // --- Slide operations ---

        if (cmd.action === "duplicate_slide") {
            // Copy slide by getting its content
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            const slideIndex = (p.slide || 1) - 1;
            // Office JS doesn't have direct duplicate - add new slide instead
            presentation.slides.add();
            await context.sync();
            return "New slide added (duplicate not directly supported in Office JS, copy content manually).";
        }

        if (cmd.action === "move_shape") {
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            const shape = shapes.items[(p.shape_index || 1) - 1];
            if (p.left !== undefined) shape.left = p.left;
            if (p.top !== undefined) shape.top = p.top;
            if (p.width !== undefined) shape.width = p.width;
            if (p.height !== undefined) shape.height = p.height;
            await context.sync();
            return "Shape moved/resized.";
        }

        if (cmd.action === "delete_shape") {
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            const shape = shapes.items[(p.shape_index || 1) - 1];
            shape.delete();
            await context.sync();
            return "Shape deleted.";
        }

        if (cmd.action === "get_shapes") {
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            let info = "Shapes on slide " + (p.slide || "last") + ":\n";
            for (let i = 0; i < shapes.items.length; i++) {
                const s = shapes.items[i];
                s.load("name,shapeType,left,top,width,height");
                await context.sync();
                let text = "";
                let fontInfo = "";
                try {
                    s.textFrame.load("textRange");
                    await context.sync();
                    s.textFrame.textRange.load("text");
                    s.textFrame.textRange.font.load("name,size,color,bold,italic");
                    await context.sync();
                    text = s.textFrame.textRange.text;
                    const f = s.textFrame.textRange.font;
                    fontInfo = " font=(" + (f.name||"?") + " " + (f.size||"?") + "pt";
                    if (f.bold) fontInfo += " bold";
                    if (f.italic) fontInfo += " italic";
                    if (f.color) fontInfo += " " + f.color;
                    fontInfo += ")";
                } catch(e) {}
                let fillInfo = "";
                try {
                    s.fill.load("foregroundColor,type");
                    await context.sync();
                    if (s.fill.type === "Solid") {
                        fillInfo = " fill=" + s.fill.foregroundColor;
                    }
                } catch(e) {}
                info += (i+1) + ". " + s.name + " (" + s.shapeType + ") " +
                        "pos=(" + Math.round(s.left) + "," + Math.round(s.top) + ") " +
                        "size=(" + Math.round(s.width) + "x" + Math.round(s.height) + ")";
                if (text) info += ' text="' + text.substring(0, 50) + '"';
                info += fontInfo + fillInfo;
                info += "\n";
            }
            return info;
        }

        // --- Add line/connector ---

        if (cmd.action === "add_line") {
            const slide = await getSlide(p.slide);
            const lineType = p.line_type || "Straight"; // Straight, StraightConnector, CurvedConnector, ElbowConnector
            const shape = slide.shapes.addLine(lineType, {
                left: p.left || 50,
                top: p.top || 50,
                width: p.width || 200,
                height: p.height || 0,
            });
            if (p.color) shape.lineFormat.color = p.color;
            if (p.weight) shape.lineFormat.weight = p.weight;
            if (p.dash_style) shape.lineFormat.dashStyle = p.dash_style; // Solid, Dash, DashDot, etc.
            // Arrow heads
            if (p.begin_arrow) shape.lineFormat.beginArrowheadStyle = p.begin_arrow; // None, Triangle, Stealth, etc.
            if (p.end_arrow) shape.lineFormat.endArrowheadStyle = p.end_arrow;
            await context.sync();
            return "Line/connector added.";
        }

        // --- Arrow connector between shapes ---

        if (cmd.action === "add_arrow") {
            // Draw an arrow line from one position to another
            const slide = await getSlide(p.slide);
            const shape = slide.shapes.addLine(p.connector_type || "StraightConnector", {
                left: p.from_left || 50,
                top: p.from_top || 50,
                width: (p.to_left || 250) - (p.from_left || 50),
                height: (p.to_top || 50) - (p.from_top || 50),
            });
            shape.lineFormat.endArrowheadStyle = p.arrow_style || "Triangle";
            if (p.color) shape.lineFormat.color = p.color;
            if (p.weight) shape.lineFormat.weight = p.weight || 2;
            await context.sync();
            return "Arrow connector added.";
        }

        // --- Rounded rectangle (corner radius via shape type) ---

        if (cmd.action === "add_rounded_rect") {
            const slide = await getSlide(p.slide);
            const shape = slide.shapes.addGeometricShape("RoundedRectangle", {
                left: p.left || 50,
                top: p.top || 50,
                width: p.width || 200,
                height: p.height || 100,
            });
            if (p.fill) shape.fill.setSolidColor(p.fill);
            if (p.no_line) shape.lineFormat.visible = false;
            if (p.line_color) shape.lineFormat.color = p.line_color;
            if (p.text) {
                shape.textFrame.textRange.text = p.text;
                if (p.font_size) shape.textFrame.textRange.font.size = p.font_size;
                if (p.font_color) shape.textFrame.textRange.font.color = p.font_color;
                if (p.font_bold) shape.textFrame.textRange.font.bold = true;
                if (p.font_name) shape.textFrame.textRange.font.name = p.font_name;
            }
            // Adjust corner radius via shape adjustment
            // adjustments[0] controls roundness (0=square, 50000=fully round)
            try {
                shape.load("adjustments");
                await context.sync();
                if (shape.adjustments && shape.adjustments.count > 0) {
                    shape.adjustments.getItemAt(0).value = p.radius || 15000; // default medium round
                }
            } catch(e) { /* adjustments may not be available */ }
            await context.sync();
            return "Rounded rectangle added.";
        }

        // --- Group shapes ---

        if (cmd.action === "group_shapes") {
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            // Select shapes by indices (1-based)
            const indices = p.shape_indices || [];
            const shapeNames = [];
            for (const idx of indices) {
                if (idx >= 1 && idx <= shapes.items.length) {
                    shapes.items[idx - 1].load("name");
                    await context.sync();
                    shapeNames.push(shapes.items[idx - 1].name);
                }
            }
            if (shapeNames.length < 2) return "Error: Need at least 2 shapes to group.";
            // Office JS doesn't have direct group API — return instruction
            return "Group not directly available in Office JS API. Shapes identified: " + shapeNames.join(", ") + ". Use manual grouping or create a compound shape instead.";
        }

        // --- Set font on specific text range within a shape ---

        if (cmd.action === "set_partial_text_style") {
            // Style a portion of text within a shape
            const slide = await getSlide(p.slide);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            const shape = shapes.items[(p.shape_index || 1) - 1];
            const textRange = shape.textFrame.textRange;
            textRange.load("text");
            await context.sync();
            // Get substring range
            const subRange = textRange.getSubstring(p.start || 0, p.length || textRange.text.length);
            if (p.font_name) subRange.font.name = p.font_name;
            if (p.font_size) subRange.font.size = p.font_size;
            if (p.font_color) subRange.font.color = p.font_color;
            if (p.bold !== undefined) subRange.font.bold = p.bold;
            if (p.italic !== undefined) subRange.font.italic = p.italic;
            await context.sync();
            return "Partial text style applied.";
        }

        // --- Apply slide layout ---

        if (cmd.action === "apply_layout") {
            const slides = presentation.slides;
            slides.load("items");
            const masters = presentation.slideMasters;
            masters.load("items");
            await context.sync();
            const slideIndex = (p.slide || 1) - 1;
            const slide = slides.items[slideIndex];
            // Get first master's layouts
            const layouts = masters.items[0].layouts;
            layouts.load("items");
            await context.sync();
            // Find layout by name or index
            let targetLayout = null;
            if (p.layout_name) {
                for (const layout of layouts.items) {
                    layout.load("name");
                    await context.sync();
                    if (layout.name.toLowerCase().includes(p.layout_name.toLowerCase())) {
                        targetLayout = layout;
                        break;
                    }
                }
            }
            if (!targetLayout && p.layout_index !== undefined) {
                targetLayout = layouts.items[p.layout_index];
            }
            if (!targetLayout) {
                // List available layouts
                let names = [];
                for (const layout of layouts.items) {
                    layout.load("name");
                    await context.sync();
                    names.push(layout.name);
                }
                return "Layout not found. Available: " + names.join(", ");
            }
            slide.layout = targetLayout;
            await context.sync();
            return "Layout applied.";
        }

        // --- List available layouts ---

        if (cmd.action === "get_layouts") {
            const masters = presentation.slideMasters;
            masters.load("items");
            await context.sync();
            const layouts = masters.items[0].layouts;
            layouts.load("items");
            await context.sync();
            let names = [];
            for (let i = 0; i < layouts.items.length; i++) {
                layouts.items[i].load("name");
                await context.sync();
                names.push((i) + ": " + layouts.items[i].name);
            }
            return "Available layouts:\n" + names.join("\n");
        }

        if (cmd.action === "set_slide_text") {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            const slideIndex = (p.slide || 1) - 1;
            const slide = slides.items[slideIndex];
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();
            for (const shape of shapes.items) {
                if (shape.textFrame) {
                    shape.textFrame.load("textRange");
                    await context.sync();
                    if (shape.textFrame.textRange) {
                        shape.textFrame.textRange.text = p.text || "";
                        await context.sync();
                        return "Text set on slide " + (slideIndex + 1);
                    }
                }
            }
            return "No text shape found on slide " + (slideIndex + 1);
        }

        // --- verify_slide_visual (alias for get_slide_image) ---

        if (cmd.action === "verify_slide_visual") {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            const slideIndex = (p.slide || 1) - 1;
            const slide = slides.items[slideIndex];
            const imgResult = slide.getImageAsBase64({ width: p.width || 960, height: p.height || 540 });
            await context.sync();
            return "__IMAGE__" + imgResult.value;
        }

        // --- list_slide_shapes (enhanced get_shapes with id, font, fill) ---

        if (cmd.action === "list_slide_shapes") {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            const slideIndex = (p.slide || 1) - 1;
            const slide = slides.items[slideIndex];
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();

            const result = [];
            for (let i = 0; i < shapes.items.length; i++) {
                const s = shapes.items[i];
                s.load("id,name,shapeType,left,top,width,height");
                await context.sync();

                const info = {
                    index: i + 1,
                    id: s.id,
                    name: s.name,
                    type: s.shapeType,
                    left: Math.round(s.left),
                    top: Math.round(s.top),
                    width: Math.round(s.width),
                    height: Math.round(s.height),
                };

                try {
                    s.textFrame.textRange.load("text");
                    s.textFrame.textRange.font.load("name,size,color,bold,italic");
                    await context.sync();
                    if (s.textFrame.textRange.text) {
                        info.text = s.textFrame.textRange.text.substring(0, 200);
                        const f = s.textFrame.textRange.font;
                        info.font = { name: f.name, size: f.size, color: f.color, bold: f.bold, italic: f.italic };
                    }
                } catch(e) {}

                try {
                    s.fill.load("foregroundColor,type");
                    await context.sync();
                    if (s.fill.type && s.fill.type !== "None") {
                        info.fill = { type: s.fill.type, color: s.fill.foregroundColor };
                    }
                } catch(e) {}

                result.push(info);
            }
            return JSON.stringify(result, null, 2);
        }

        // --- set_z_order ---

        if (cmd.action === "set_z_order") {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();
            const slideIndex = (p.slide || 1) - 1;
            const slide = slides.items[slideIndex];
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();

            let targetShape = null;
            if (p.shape_id !== undefined) {
                for (const s of shapes.items) {
                    s.load("id");
                    await context.sync();
                    if (String(s.id) === String(p.shape_id)) { targetShape = s; break; }
                }
            } else if (p.shape_index !== undefined) {
                targetShape = shapes.items[(p.shape_index || 1) - 1];
            }

            if (!targetShape) return "Error: Shape not found. Specify shape_id or shape_index.";

            const orderMap = {
                "BringToFront": Office.ZOrderType.BringToFront,
                "SendToBack":   Office.ZOrderType.SendToBack,
                "BringForward": Office.ZOrderType.BringForward,
                "SendBackward": Office.ZOrderType.SendBackward,
            };
            const zOrderType = orderMap[p.order];
            if (zOrderType === undefined)
                return "Error: Invalid order. Use BringToFront, SendToBack, BringForward, or SendBackward.";

            targetShape.setZOrder(zOrderType);
            await context.sync();
            return "Z-order set to " + p.order + ".";
        }

        // --- verify_slides ---

        if (cmd.action === "verify_slides") {
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();

            const fromIdx = (p.from_slide || 1) - 1;
            const toIdx   = Math.min((p.to_slide || slides.items.length) - 1, slides.items.length - 1);
            const SLIDE_W = 960, SLIDE_H = 540, MARGIN = 30;
            const results = [];

            for (let si = fromIdx; si <= toIdx; si++) {
                const slide = slides.items[si];
                const shapes = slide.shapes;
                shapes.load("items");
                await context.sync();

                const shapeData = [];
                for (const s of shapes.items) {
                    s.load("name,shapeType,left,top,width,height");
                    await context.sync();
                    let text = '', textColor = null, fillColor = null;
                    try {
                        s.textFrame.textRange.load("text");
                        s.textFrame.textRange.font.load("color");
                        await context.sync();
                        text = s.textFrame.textRange.text || '';
                        textColor = s.textFrame.textRange.font.color;
                    } catch(e) {}
                    try {
                        s.fill.load("foregroundColor,type");
                        await context.sync();
                        if (s.fill.type === "Solid") fillColor = s.fill.foregroundColor;
                    } catch(e) {}
                    shapeData.push({ name: s.name, left: s.left, top: s.top, width: s.width, height: s.height, text, textColor, fillColor });
                }

                if (shapeData.length === 0) {
                    results.push({ slide: si+1, coverage_ratio: 0, bottom_gap: SLIDE_H, overlaps: [], out_of_bounds: [], contrast_warnings: [] });
                    continue;
                }

                const minX = Math.min(...shapeData.map(s => s.left));
                const minY = Math.min(...shapeData.map(s => s.top));
                const maxX = Math.max(...shapeData.map(s => s.left + s.width));
                const maxY = Math.max(...shapeData.map(s => s.top + s.height));

                const safeArea = (SLIDE_W - 2*MARGIN) * (SLIDE_H - 2*MARGIN);
                const coverage_ratio = parseFloat(Math.min(1, ((maxX-minX)*(maxY-minY)) / safeArea).toFixed(3));
                const bottom_gap = Math.round(SLIDE_H - maxY);

                // Overlapping shape pairs
                const overlaps = [];
                for (let i = 0; i < shapeData.length; i++) {
                    for (let j = i+1; j < shapeData.length; j++) {
                        const a = shapeData[i], b = shapeData[j];
                        if (a.left < b.left+b.width && a.left+a.width > b.left &&
                            a.top < b.top+b.height && a.top+a.height > b.top) {
                            overlaps.push([a.name, b.name]);
                        }
                    }
                }

                // Out of slide bounds
                const out_of_bounds = shapeData
                    .filter(s => s.left < 0 || s.top < 0 || s.left+s.width > SLIDE_W || s.top+s.height > SLIDE_H)
                    .map(s => s.name);

                // Contrast warnings (text vs fill)
                const contrast_warnings = [];
                for (const s of shapeData) {
                    if (s.text && s.textColor && s.fillColor) {
                        const ratio = getContrastRatio(s.textColor, s.fillColor);
                        if (ratio < 4.5) {
                            contrast_warnings.push({ shape: s.name, ratio: parseFloat(ratio.toFixed(2)) });
                        }
                    }
                }

                results.push({ slide: si+1, coverage_ratio, bottom_gap, overlaps, out_of_bounds, contrast_warnings });
            }
            return JSON.stringify(results, null, 2);
        }

        return "Unknown PowerPoint action: " + cmd.action;
    });
}
