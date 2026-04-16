---
description: Create and edit PowerPoint presentations with full OOXML and JSZip support via ClaudeGateway
when_to_use: 사용자가 PowerPoint 슬라이드 생성, 편집, 분석을 요청할 때
---

# PowerPoint Pack

You can create, edit, and analyze PowerPoint presentations through the Office Add-in bridge.
The Add-in must be loaded in PowerPoint before using any tools.

---

## Tools

| Tool | When to use |
|------|-------------|
| `ppt_list_shapes(slide)` | Read all shapes with id, position, size, text, fill — **always call before modifying** |
| `ppt_screenshot(slide)` | Screenshot a slide — use to visually confirm every major edit |
| `ppt_verify(from_slide?, to_slide?)` | Check coverage_ratio, bottom_gap, overlaps, contrast — call after each slide is done |
| `ppt_edit_xml(slide, code)` | **Primary tool.** Edit slide OOXML via JSZip JavaScript code — use for precise multi-element layouts, backgrounds |
| `ppt_edit_chart(slide, code)` | Insert/replace a chart using OOXML 3-file pattern (chart XML + Content_Types + rels) |
| `ppt_set_z_order(slide, shape_id, order)` | Fix z-order: BringToFront / SendToBack / BringForward / SendBackward |
| `ppt_command(action, params)` | All other Office.js ops: add_slide, add_text_box, add_shape, set_background, delete_slide, get_slide_count, get_layouts, etc. |
| `ppt_get_context()` | Check if Office Add-in is connected and get PowerPoint document state |

---

## Coordinate System

- Slide size: **960 × 540 pt**
- Unit conversion: **1 pt = 12,700 EMU**
- Always work in pt, convert to EMU only when writing OOXML

```
EMU = pt × 12700
pt  = EMU ÷ 12700

# Common card layout formula (n cards across)
card_width = (960 - left_margin - right_margin - gap × (n-1)) / n
```

---

## Core Principles

1. **Reuse existing slides.** Always call `ppt_get_context` first. If slides already exist, modify them — do NOT delete and recreate. Only add new slides when the existing count is insufficient.
2. **One slide at a time.** Complete each slide fully before moving to the next. Take screenshot, verify, fix — then proceed.
2. **Read before you write.** Always call `ppt_list_shapes()` before modifying an existing slide.
3. **Prefer OOXML for complex layouts.** `ppt_edit_xml` places multiple elements in one shot — faster and more precise than multiple `ppt_command` calls.
4. **Charts require OOXML.** `ppt_command` has no chart creation API. Always use `ppt_edit_chart`.
5. **Verify after every edit.** Call `ppt_screenshot` immediately after each slide edit. Fix issues before moving on.
6. **Use shape IDs, not names.** Shape names are locale-dependent and unreliable. Always use IDs from `ppt_list_shapes`.
7. **Blank layout for custom designs.** Use "Blank" layout (no placeholders) when you control all element placement.

---

## Workflow Patterns

### Starting any task

```
1. ppt_get_context()                ← check how many slides exist
2. If slides exist → modify them (not delete)
3. If more slides needed → ppt_command("add_slide")
```

### Working on one slide

```
1. ppt_edit_xml(slide, code)        ← place all elements via OOXML in one shot
2. ppt_screenshot(slide)            ← visual confirmation (ONE slide only, not all at once)
3. Fix issues if found, repeat from step 1
4. Only after this slide is done → move to next slide
```

### Editing an existing slide

```
1. ppt_list_shapes(slide)           ← read current state (use id, not name)
2. ppt_screenshot(slide)            ← see what it looks like now
3. ppt_edit_xml(slide, code)        ← make OOXML changes
4. ppt_verify(slide)
5. ppt_screenshot(slide)            ← confirm result
```

### Adding a chart

```
1. ppt_command("add_slide")
2. ppt_edit_chart(slide, code)      ← code modifies 3 files in zip:
                                       ppt/charts/chart{n}.xml (c:chartSpace)
                                       [Content_Types].xml (Override entry)
                                       ppt/slides/_rels/slide{n}.xml.rels (relationship)
3. ppt_verify(slide)
4. ppt_screenshot(slide)
```

### Creating a diagram / flowchart

```
1. ppt_command("add_slide")
2. ppt_edit_xml(slide, code)        ← all boxes, arrows, labels in one OOXML shot
3. ppt_screenshot(slide)            ← check z-order and alignment
4. ppt_set_z_order(slide, shape_id, "BringToFront")  ← fix layering if needed
5. ppt_verify(slide)
```

---

<!-- examples -->

## ppt_edit_xml Code Pattern

**ALWAYS use full slide XML replacement. Never try to parse/modify existing XML.**

```javascript
// code parameter for ppt_edit_xml — full slide replacement pattern
const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="1E3A5F"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="9144000" cy="5143500"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="9144000" cy="5143500"/>
        </a:xfrm>
      </p:grpSpPr>
      <!-- INSERT SHAPES HERE -->
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;

zip.file(slidePath, slideXml);
markDirty(); // ← 반드시 호출! 없으면 변경사항이 PowerPoint에 적용되지 않음
```

**Rules — MUST FOLLOW ALL:**
- **마지막 줄은 반드시 `markDirty()`** — 없으면 변경사항이 PowerPoint에 적용되지 않음. 절대 빠뜨리지 말 것.
- **JSZip 3.x API**: read → `await zip.files[slidePath].async('string')` (NOT `.asText()`)
- Write → `zip.file(slidePath, xmlString)` 그 다음 반드시 `markDirty()`
- Do NOT use DOMParser, XMLSerializer, or `.children` — XML을 문자열로 직접 작성
- Escape: `&` → `&amp;`  `<` → `&lt;`  `>` → `&gt;`
- Slide size: cx="9144000" cy="5143500" (EMU) = 960×540 pt
- Every shape needs unique numeric `id` (1, 2, 3, ...)

---

## OOXML Quick Reference

### Shape with text
```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="1" name="Box1"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1143000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="1F497D"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="ko-KR" sz="1800" b="1"/><a:t>Label</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
```

### Arrow connector
```xml
<p:sp>
  <p:spPr>
    <a:xfrm><a:off x="3657600" y="1371600"/><a:ext cx="457200" cy="0"/></a:xfrm>
    <a:prstGeom prst="notchedRightArrow"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
  </p:spPr>
</p:sp>
```

### Slide background
```xml
<p:bg>
  <p:bgPr>
    <a:solidFill><a:srgbClr val="1E3A5F"/></a:solidFill>
    <a:effectLst/>
  </p:bgPr>
</p:bg>
```

---

## Design Defaults

When no specific design is requested, use these defaults:

```
Background:     #1E3A5F (dark navy)  or  #FFFFFF (white)
Primary color:  #4472C4 (blue)
Accent:         #ED7D31 (orange)
Text (dark bg): #FFFFFF
Text (light bg): #1F1F1F
Font (Korean):  "맑은 고딕" or "Noto Sans KR"
Font (English): "Calibri" or "Arial"
Title size:     32–40pt
Body size:      16–24pt
Min font size:  14pt (never go below)
Slide margin:   30pt on all sides
```

---

## Common Mistakes to Avoid

- **Never simulate charts with shapes.** Use `ppt_edit_chart` for all charts.
- **Never use shape names to find shapes.** They change by locale. Use IDs.
- **Never skip verification.** OOXML coordinates can look correct but render wrongly.
- **Never place text below 14pt.** It becomes unreadable in presentations.
- **Do not leave large blank areas.** If bottom gap > 80pt, expand content or reposition.
- **Do not use emojis in slide content.** Rendering is unstable across PowerPoint versions.
