import {
  XmlHelper
} from "./chunk-JB6TXILZ.js";

// src/core/PptxReader.ts
import JSZip from "jszip";

// src/elements/TextExtractor.ts
var TextExtractor = class {
  /**
   * Extracts all text elements from the slide's <spTree> node.
   * @param spTree The <spTree> element of the slide.
   * @param themeColors Theme color mapping (e.g. tx1, bg2).
   * @returns List of TextElement extracted.
   */
  static extract(spTree, themeColors, opts = {}) {
    if (!spTree) return [];
    const elements = [];
    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const nvPr = shape.getElementsByTagNameNS("*", "nvPr")[0] ?? null;
      const ph = nvPr?.getElementsByTagNameNS("*", "ph")[0] ?? null;
      const isPlaceholder = !!ph;
      if (opts.context && opts.context !== "slide" && isPlaceholder) {
        continue;
      }
      const txBody = shape.getElementsByTagNameNS("*", "txBody")[0];
      if (!txBody) continue;
      const paragraphs = txBody.getElementsByTagNameNS("*", "p");
      const bodyPr = txBody.getElementsByTagNameNS("*", "bodyPr")[0] ?? null;
      const anchor = bodyPr?.getAttribute("anchor") || void 0;
      const verticalAlign = anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top";
      const lIns = bodyPr?.getAttribute("lIns");
      const tIns = bodyPr?.getAttribute("tIns");
      const rIns = bodyPr?.getAttribute("rIns");
      const bIns = bodyPr?.getAttribute("bIns");
      const padding = {
        left: lIns ? Number(lIns) / 9525 : 0,
        top: tIns ? Number(tIns) / 9525 : 0,
        right: rIns ? Number(rIns) / 9525 : 0,
        bottom: bIns ? Number(bIns) / 9525 : 0
      };
      const textRuns = [];
      let fontName = "Arial";
      let fontSize = 18;
      let color = void 0;
      let horizontalAlign = void 0;
      const paraItems = [];
      const lvlDefaults = {};
      const lstStyle = txBody.querySelector("*|lstStyle");
      if (lstStyle) {
        for (const node of Array.from(lstStyle.children)) {
          const m = node.localName.match(/^lvl(\d+)pPr$/);
          if (!m) continue;
          const idx = parseInt(m[1], 10) - 1;
          let kind = "p";
          let listStyle = void 0;
          if (node.querySelector("*|buNone")) {
            kind = "p";
          } else if (node.querySelector("*|buAutoNum")) {
            kind = "ol";
            const auto = node.querySelector("*|buAutoNum");
            const typ = auto?.getAttribute("type") || "arabicPeriod";
            listStyle = mapAutoNumToCss(typ);
          } else if (node.querySelector("*|buChar")) {
            kind = "ul";
            listStyle = "disc";
          }
          lvlDefaults[idx] = { kind, listStyle };
        }
      }
      for (const p of Array.from(paragraphs)) {
        const pPr = p.getElementsByTagNameNS("*", "pPr")[0] ?? null;
        const algn = pPr?.getAttribute("algn") || void 0;
        if (algn && !horizontalAlign) {
          horizontalAlign = algn === "ctr" ? "center" : algn === "r" ? "right" : algn.startsWith("just") ? "justify" : "left";
        }
        const runs = p.getElementsByTagNameNS("*", "r");
        let paraText = getParagraphText(p);
        const paraHtml = getParagraphHtml(p, themeColors);
        if (paraText) paraText.split(/\n+/).forEach((t) => {
          if (t) textRuns.push(t);
        });
        for (const r of Array.from(runs)) {
          const rPr = r.getElementsByTagNameNS("*", "rPr")[0];
          if (rPr) {
            const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
            fontName = latin?.getAttribute("typeface") ?? fontName;
            const sz = rPr.getAttribute("sz");
            if (sz) {
              const n = parseInt(sz, 10);
              if (Number.isFinite(n)) fontSize = n / 100;
            }
            const solidFill = rPr.querySelector("*|solidFill");
            const candidate = XmlHelper.getColorFromElement(solidFill || null, themeColors);
            if (candidate) color = candidate;
          }
        }
        if (!color && pPr) {
          const endParaRPr = pPr.getElementsByTagNameNS("*", "endParaRPr")[0] ?? null;
          const endFill = endParaRPr?.querySelector("*|solidFill") || null;
          const c1 = XmlHelper.getColorFromElement(endFill, themeColors);
          if (c1) color = c1;
          if (!color) {
            const defRPrP = pPr.getElementsByTagNameNS("*", "defRPr")[0] ?? null;
            const defFillP = defRPrP?.querySelector("*|solidFill") || null;
            const c2 = XmlHelper.getColorFromElement(defFillP, themeColors);
            if (c2) color = c2;
          }
        }
        const lvlAttr = pPr?.getAttribute("lvl");
        const lvl = lvlAttr ? parseInt(lvlAttr, 10) : 0;
        let kind = "p";
        let listStyle = void 0;
        if (pPr) {
          if (pPr.querySelector("*|buNone")) {
            kind = "p";
          } else if (pPr.querySelector("*|buAutoNum")) {
            kind = "ol";
            const auto = pPr.querySelector("*|buAutoNum");
            const typ = auto?.getAttribute("type") || "arabicPeriod";
            listStyle = mapAutoNumToCss(typ);
          } else if (pPr.querySelector("*|buChar")) {
            kind = "ul";
            listStyle = "disc";
          } else if (lvlDefaults[lvl]) {
            kind = lvlDefaults[lvl].kind;
            listStyle = lvlDefaults[lvl].listStyle;
          } else if (lvl > 0) {
            kind = "ul";
            listStyle = "disc";
          }
        } else if (lvlDefaults[lvl]) {
          kind = lvlDefaults[lvl].kind;
          listStyle = lvlDefaults[lvl].listStyle;
        } else if (lvl > 0) {
          kind = "ul";
          listStyle = "disc";
        }
        let lineHeightPt = void 0;
        const lnSpc = pPr?.querySelector("*|lnSpc");
        if (lnSpc) {
          const spcPtsEl = lnSpc.querySelector("*|spcPts");
          const spcPctEl = lnSpc.querySelector("*|spcPct");
          if (spcPtsEl) {
            const val = spcPtsEl.getAttribute("val");
            if (val) lineHeightPt = parseInt(val, 10) / 100;
          } else if (spcPctEl) {
            const val = spcPctEl.getAttribute("val");
            if (val) lineHeightPt = parseInt(val, 10) / 1e5 * fontSize;
          }
        }
        const spcBefVal = pPr?.querySelector("*|spcBef > *|spcPts")?.getAttribute("val");
        const spaceBeforePt = spcBefVal ? parseInt(spcBefVal, 10) / 100 : void 0;
        const spcAftVal = pPr?.querySelector("*|spcAft > *|spcPts")?.getAttribute("val");
        const spaceAfterPt = spcAftVal ? parseInt(spcAftVal, 10) / 100 : void 0;
        paraItems.push({ kind, text: paraText, html: paraHtml, lvl: isNaN(lvl) ? 0 : lvl, listStyle, lineHeightPt, spaceBeforePt, spaceAfterPt });
      }
      if (!color) {
        const lstDefRPr = txBody.querySelector("*|lstStyle *|defRPr");
        const lstFill = lstDefRPr?.querySelector("*|solidFill");
        const c0 = XmlHelper.getColorFromElement(lstFill || null, themeColors);
        if (c0) color = c0;
        const defRPr = txBody.querySelector("*|defRPr");
        const defFill = defRPr?.querySelector("*|solidFill");
        const fallback1 = XmlHelper.getColorFromElement(defFill || null, themeColors);
        if (fallback1) color = fallback1;
        if (!color) {
          const spPr = shape.querySelector("p\\:spPr, spPr");
          const shapeFill = spPr?.querySelector("*|solidFill");
          const fallback2 = XmlHelper.getColorFromElement(shapeFill || null, themeColors);
          if (fallback2) color = fallback2;
        }
      }
      const content = textRuns.join(" ").trim();
      if (opts.context && opts.context !== "slide") {
        const c = content.toLowerCase();
        const isDefault = c.includes("click to add") || c.includes("click to edit") || c.includes("haga clic para agregar") || c.includes("haga clic para editar") || c.includes("hacer clic para agregar") || c.includes("hacer clic para editar");
        if (isDefault) continue;
      }
      if (content === "") continue;
      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
      let off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      let ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;
      let x, y, cx, cy;
      if (off && ext) {
        x = XmlHelper.getAttrAsNumber(off, "x");
        y = XmlHelper.getAttrAsNumber(off, "y");
        cx = XmlHelper.getAttrAsNumber(ext, "cx");
        cy = XmlHelper.getAttrAsNumber(ext, "cy");
      } else if (opts.placeholderGeom) {
        const phIdx = ph?.getAttribute("idx") || void 0;
        const phType = ph?.getAttribute("type") || void 0;
        const g = (phIdx ? opts.placeholderGeom[phIdx] : void 0) || (phType ? opts.placeholderGeom[phType] : void 0);
        x = g?.x ?? 0;
        y = g?.y ?? 0;
        cx = g?.cx ?? 1e6;
        cy = g?.cy ?? 5e5;
        if (g?.fontSize && fontSize === 18) {
          fontSize = g.fontSize;
        }
      } else {
        x = 0;
        y = 0;
        cx = 1e6;
        cy = 5e5;
      }
      let richHtml = void 0;
      {
        const parts = [];
        let open = null;
        for (const it of paraItems) {
          const spacingCss = [
            it.lineHeightPt ? `line-height:${it.lineHeightPt}pt` : "",
            it.spaceBeforePt ? `margin-top:${it.spaceBeforePt}pt` : "",
            it.spaceAfterPt ? `margin-bottom:${it.spaceAfterPt}pt` : ""
          ].filter(Boolean).join(";");
          if (it.kind === "p") {
            if (open) {
              parts.push(open.kind === "ul" ? "</ul>" : "</ol>");
              open = null;
            }
            if (it.text.trim()) {
              parts.push(`<div style="margin-left:${it.lvl * 24}px${spacingCss ? `;${spacingCss}` : ""}">${it.html}</div>`);
            }
            continue;
          }
          if (!open || open.kind !== it.kind) {
            if (open) parts.push(open.kind === "ul" ? "</ul>" : "</ol>");
            const commonListCss = `list-style-position: inside; padding-left: 0; margin: 0;`;
            const style = it.kind === "ol" ? ` style="${commonListCss} list-style-type: ${it.listStyle || "decimal"};"` : ` style="${commonListCss}"`;
            parts.push(it.kind === "ul" ? `<ul${style}>` : `<ol${style}>`);
            open = { kind: it.kind, listStyle: it.listStyle };
          }
          parts.push(`<li style="margin-left:${it.lvl * 24}px${spacingCss ? `;${spacingCss}` : ""}">${it.html}</li>`);
        }
        if (open) parts.push(open.kind === "ul" ? "</ul>" : "</ol>");
        richHtml = parts.join("");
      }
      const element = {
        type: "text",
        content,
        position: { x, y },
        size: { width: cx, height: cy },
        font: {
          name: fontName,
          size: fontSize,
          color: color ?? "#000000"
          // fallback absoluto si quieres
        },
        align: {
          horizontal: horizontalAlign ?? "left",
          vertical: verticalAlign
        },
        padding,
        html: richHtml
      };
      elements.push(element);
    }
    return elements;
  }
};
function getParagraphText(p) {
  let out = "";
  for (const child of Array.from(p.childNodes)) {
    if (!(child instanceof Element)) {
      continue;
    }
    const ln = child.localName;
    if (ln === "r") {
      const t = child.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
      out += t;
    } else if (ln === "br") {
      out += "\n";
    } else if (ln === "fld") {
      const runs = child.getElementsByTagNameNS("*", "r");
      for (const r of Array.from(runs)) {
        const t = r.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
        out += t;
      }
    } else if (ln === "tab") {
      out += "	";
    }
  }
  return out;
}
function getParagraphHtml(p, themeColors) {
  let out = "";
  for (const child of Array.from(p.childNodes)) {
    if (!(child instanceof Element)) continue;
    const ln = child.localName;
    if (ln === "r") {
      const t = child.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
      if (!t) continue;
      const rPr = child.getElementsByTagNameNS("*", "rPr")[0] ?? null;
      const styles = [];
      if (rPr) {
        const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
        const typeface = latin?.getAttribute("typeface");
        if (typeface) styles.push(`font-family:${typeface}, sans-serif`);
        const sz = rPr.getAttribute("sz");
        if (sz) {
          const n = parseInt(sz, 10);
          if (Number.isFinite(n)) styles.push(`font-size:${n / 100}pt`);
        }
        const solidFill = rPr.querySelector("*|solidFill");
        const color = XmlHelper.getColorFromElement(solidFill || null, themeColors);
        if (color) styles.push(`color:${color}`);
        if (rPr.getAttribute("b") === "1") styles.push("font-weight:bold");
        if (rPr.getAttribute("i") === "1") styles.push("font-style:italic");
        const u = rPr.getAttribute("u");
        if (u && u !== "none") styles.push("text-decoration:underline");
      }
      const escaped = escapeHtml(t);
      out += styles.length > 0 ? `<span style="${styles.join(";")}">${escaped}</span>` : escaped;
    } else if (ln === "br") {
      out += "<br>";
    } else if (ln === "fld") {
      const fldRuns = child.getElementsByTagNameNS("*", "r");
      if (fldRuns.length) {
        for (const r of Array.from(fldRuns)) {
          const t = r.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
          out += escapeHtml(t);
        }
      } else {
        const t = child.getElementsByTagNameNS("*", "t")[0]?.textContent;
        if (t) out += escapeHtml(t);
      }
    } else if (ln === "tab") {
      out += "&emsp;";
    }
  }
  return out;
}
function mapAutoNumToCss(typ) {
  const t = typ.toLowerCase();
  if (t.includes("alphauc")) return "upper-alpha";
  if (t.includes("alphalc")) return "lower-alpha";
  if (t.includes("romanu")) return "upper-roman";
  if (t.includes("romanl")) return "lower-roman";
  return "decimal";
}
function escapeHtml(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

// src/elements/ImageExtractor.ts
var ImageExtractor = class {
  /**
   * Extracts image elements from the <spTree> element using rels from slide relationships.
   * @param spTree The <spTree> element of the slide.
   * @param rels XML Document for slide relationships (ppt/slides/_rels/slideX.xml.rels).
   * @param zip The JSZip archive of the entire .pptx file.
   * @returns List of ImageElement extracted.
   */
  static async extract(spTree, rels, zip, basePath = "ppt/slides") {
    if (!spTree) return [];
    const elements = [];
    const pics = spTree.getElementsByTagNameNS("*", "pic");
    for (const pic of Array.from(pics)) {
      const blip = pic.getElementsByTagNameNS("*", "blip")[0];
      const embedId = blip?.getAttribute("r:embed") ?? "";
      if (!embedId) continue;
      const relEl = rels && rels.getElementsByTagName ? function() {
        const els = rels.getElementsByTagName("Relationship");
        for (const e of Array.from(els)) {
          if (e.getAttribute("Id") === embedId) return e;
        }
        return null;
      }() : null;
      const relTarget = relEl?.getAttribute("Target");
      if (!relTarget) continue;
      const normalizedPath = this.normalizePath(relTarget, basePath);
      const imageFile = zip.file(normalizedPath);
      if (!imageFile) continue;
      const binary = await imageFile.async("base64");
      const extImg = normalizedPath.split(".").pop()?.toLowerCase() || "png";
      const dataUri = `data:image/${extImg};base64,${binary}`;
      const xfrm = pic.getElementsByTagNameNS("*", "xfrm")[0];
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];
      const x = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
      const y = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;
      const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1e6;
      const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 5e5;
      const element = {
        type: "image",
        relId: embedId,
        src: dataUri,
        position: { x, y },
        size: { width: cx, height: cy }
      };
      elements.push(element);
    }
    return elements;
  }
  /**
   * Normalizes a relative path from a slide rels file.
   * @param target Path from the relationship XML (e.g. "../media/image1.png")
   * @param basePath Base folder (e.g. "ppt/slides")
   * @returns Normalized path inside the zip (e.g. "ppt/media/image1.png")
   */
  static normalizePath(target, basePath) {
    const parts = (basePath + "/" + target).split("/");
    const resolved = [];
    for (const part of parts) {
      if (part === "..") resolved.pop();
      else if (part !== ".") resolved.push(part);
    }
    return resolved.join("/");
  }
};

// src/elements/ShapeExtractor.ts
var ShapeExtractor = class {
  /**
   * Extracts shape and connector elements from the <spTree> element of the slide.
   * @param spTree The <spTree> element.
   * @param themeColors Theme color mapping.
   * @returns List of ShapeElement extracted.
   */
  static extract(spTree, themeColors) {
    if (!spTree) return [];
    const elements = [];
    const allShapes = [
      ...Array.from(spTree.getElementsByTagNameNS("*", "sp")),
      ...Array.from(spTree.getElementsByTagNameNS("*", "cxnSp"))
    ];
    for (const shape of allShapes) {
      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];
      const x = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
      const y = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;
      const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1e6;
      const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 5e5;
      const rotAttr = xfrm?.getAttribute("rot");
      const rotationDeg = rotAttr ? Number(rotAttr) / 6e4 : void 0;
      const flipH = xfrm?.getAttribute("flipH") === "1";
      const flipV = xfrm?.getAttribute("flipV") === "1";
      const prstGeom = shape.getElementsByTagNameNS("*", "prstGeom")[0];
      const shapeType = prstGeom?.getAttribute("prst") ?? "rect";
      let cornerRadiusPct = void 0;
      if (shapeType === "roundRect" && prstGeom) {
        const avLst = prstGeom.getElementsByTagNameNS("*", "avLst")[0];
        const gd = avLst?.getElementsByTagNameNS("*", "gd")[0];
        if (gd?.getAttribute("name") === "adj") {
          const fmla = gd.getAttribute("fmla") ?? "";
          const match = fmla.match(/val\s+(\d+)/);
          if (match) cornerRadiusPct = parseInt(match[1], 10);
        }
      }
      const spPr = shape.getElementsByTagNameNS("*", "spPr")[0];
      const fill = this.extractFill(spPr, shape, themeColors);
      const stroke = this.extractStroke(spPr, themeColors);
      const element = {
        type: "shape",
        shapeType,
        position: { x, y },
        size: { width: cx, height: cy },
        fill,
        stroke,
        rotationDeg,
        flipH: flipH || void 0,
        flipV: flipV || void 0,
        cornerRadiusPct
      };
      elements.push(element);
    }
    return elements;
  }
  /** Extract fill from shape properties, with theme/style fallback */
  static extractFill(spPr, shape, themeColors) {
    if (spPr) {
      for (const child of Array.from(spPr.children)) {
        const tag = child.localName;
        if (tag === "noFill") {
          return { type: "none" };
        }
        if (tag === "solidFill") {
          const color = XmlHelper.getColorFromElement(child, themeColors);
          if (color) return { type: "solid", color };
        }
        if (tag === "gradFill") {
          const gradient = XmlHelper.getGradientFromElement(child, themeColors);
          if (gradient) return gradient;
        }
      }
    }
    const style = shape.getElementsByTagNameNS("*", "style")[0];
    const fillRef = style?.getElementsByTagNameNS("*", "fillRef")[0];
    const schemeClr = fillRef?.getElementsByTagNameNS("*", "schemeClr")[0];
    const val = schemeClr?.getAttribute("val");
    if (val && themeColors[val]) {
      return { type: "solid", color: themeColors[val] };
    }
    return { type: "none" };
  }
  /** Extract stroke/line properties */
  static extractStroke(spPr, themeColors) {
    if (!spPr) return void 0;
    const ln = spPr.getElementsByTagNameNS("*", "ln")[0];
    if (!ln) return void 0;
    for (const child of Array.from(ln.children)) {
      if (child.localName === "noFill") return void 0;
    }
    const borderFill = ln.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
    const color = XmlHelper.getColorFromElement(borderFill, themeColors) ?? "transparent";
    const wAttr = ln.getAttribute("w");
    let width = 1;
    if (wAttr) {
      const w = Number(wAttr);
      if (!isNaN(w)) {
        width = w / 9525;
      }
    }
    const headEndEl = ln.getElementsByTagNameNS("*", "headEnd")[0] ?? null;
    const tailEndEl = ln.getElementsByTagNameNS("*", "tailEnd")[0] ?? null;
    const headEnd = headEndEl ? {
      type: headEndEl.getAttribute("type") || void 0,
      w: headEndEl.getAttribute("w") || void 0,
      len: headEndEl.getAttribute("len") || void 0
    } : void 0;
    const tailEnd = tailEndEl ? {
      type: tailEndEl.getAttribute("type") || void 0,
      w: tailEndEl.getAttribute("w") || void 0,
      len: tailEndEl.getAttribute("len") || void 0
    } : void 0;
    const prstDash = ln.getElementsByTagNameNS("*", "prstDash")[0];
    const dashStyle = prstDash?.getAttribute("val") || void 0;
    if (color === "transparent" && !headEnd && !tailEnd && !dashStyle && width <= 0) {
      return void 0;
    }
    return { color, width, dashStyle, headEnd, tailEnd };
  }
};

// src/elements/TableExtractor.ts
var TableExtractor = class {
  static extract(spTree, themeColors, themeTableStyles) {
    if (!spTree) return [];
    const tables = [];
    const gFrames = spTree.getElementsByTagNameNS("*", "graphicFrame");
    for (const gf of Array.from(gFrames)) {
      const tbl = this.findTbl(gf);
      if (!tbl) continue;
      const xfrm = gf.getElementsByTagNameNS("*", "xfrm")[0] ?? null;
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;
      const x = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
      const y = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;
      const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1e6;
      const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 5e5;
      const columns = [];
      const grid = tbl.getElementsByTagNameNS("*", "tblGrid")[0] ?? null;
      if (grid) {
        for (const col of Array.from(grid.getElementsByTagNameNS("*", "gridCol"))) {
          const w = Number(col.getAttribute("w") || 0);
          columns.push(w);
        }
      }
      const tblPr = tbl.getElementsByTagNameNS("*", "tblPr")[0] ?? null;
      const tableStyle = this.extractTableStyleFlags(tblPr);
      const tableBorders = this.extractTableBorders(tblPr, themeColors);
      const tableStyleId = this.extractTableStyleId(tblPr);
      const tableFillColor = this.extractTableFill(tblPr, themeColors);
      const resolvedStyle = tableStyleId && themeTableStyles ? themeTableStyles[tableStyleId] : void 0;
      const fallbackStyle = this.buildFallbackTableStyle(themeColors, tableStyle);
      const mergedStyle = fallbackStyle || resolvedStyle ? {
        fills: { ...fallbackStyle?.fills || {}, ...resolvedStyle?.fills || {} },
        fontColors: { ...fallbackStyle?.fontColors || {}, ...resolvedStyle?.fontColors || {} }
      } : void 0;
      const rows = [];
      for (const tr of Array.from(tbl.getElementsByTagNameNS("*", "tr"))) {
        const cells = [];
        for (const tc of Array.from(tr.getElementsByTagNameNS("*", "tc"))) {
          const txBody = tc.getElementsByTagNameNS("*", "txBody")[0] ?? null;
          const { text, font, align, padding } = this.extractCellText(txBody, themeColors, tc);
          const tcPr = tc.getElementsByTagNameNS("*", "tcPr")[0] ?? null;
          const fillColor = this.extractFillColor(tcPr, themeColors);
          const borders = this.extractCellBorders(tcPr, themeColors);
          const cell = { text, font, align, padding, fillColor, borders };
          const gridSpan = tcPr?.getElementsByTagNameNS("*", "gridSpan")[0] ?? null;
          const rowSpan = tcPr?.getElementsByTagNameNS("*", "rowSpan")[0] ?? null;
          if (gridSpan) cell.colSpan = Number(gridSpan.getAttribute("val") || 1);
          if (rowSpan) cell.rowSpan = Number(rowSpan.getAttribute("val") || 1);
          cells.push(cell);
        }
        rows.push({ cells });
      }
      tables.push({
        type: "table",
        position: { x, y },
        size: { width: cx, height: cy },
        columns,
        rows,
        tableStyle,
        tableBorders,
        tableStyleId,
        tableFillColor,
        style: mergedStyle ? { fills: mergedStyle.fills, fontColors: mergedStyle.fontColors } : void 0
      });
    }
    return tables;
  }
  static buildFallbackTableStyle(themeColors, tableStyleFlags) {
    const accent = themeColors["accent1"] || themeColors["accent2"] || themeColors["dk1"];
    if (!accent) return void 0;
    const white = themeColors["lt1"] || "#FFFFFF";
    const fills = {};
    const fontColors = {};
    if (tableStyleFlags?.firstRow) {
      fills["firstRow"] = accent;
      fontColors["firstRow"] = white;
    }
    if (tableStyleFlags?.firstCol) {
      fills["firstCol"] = this.lightenHex(accent, 0.85);
      fontColors["firstCol"] = void 0;
    }
    const bandShade = this.lightenHex(accent, 0.92);
    if (tableStyleFlags?.bandRow) {
      fills["band2H"] = bandShade;
    }
    if (tableStyleFlags?.bandCol) {
      fills["band2V"] = bandShade;
    }
    return { fills, fontColors };
  }
  static lightenHex(hex, ratio) {
    const m = (hex || "").replace("#", "");
    if (m.length !== 6 || /[^0-9a-fA-F]/.test(m)) return hex || "#FFFFFF";
    const r = parseInt(m.substring(0, 2), 16);
    const g = parseInt(m.substring(2, 4), 16);
    const b = parseInt(m.substring(4, 6), 16);
    const lr = Math.round(r + (255 - r) * ratio);
    const lg = Math.round(g + (255 - g) * ratio);
    const lb = Math.round(b + (255 - b) * ratio);
    const to2 = (n) => n.toString(16).padStart(2, "0");
    return `#${to2(lr)}${to2(lg)}${to2(lb)}`;
  }
  static findTbl(gf) {
    const graphicData = gf.getElementsByTagNameNS("*", "graphicData")[0] ?? null;
    if (!graphicData) return null;
    const tbl = graphicData.getElementsByTagNameNS("*", "tbl")[0] ?? null;
    return tbl;
  }
  static extractCellText(txBody, themeColors, tc) {
    if (!txBody) return { text: "", font: {}, align: {}, padding: { left: 0, top: 0, right: 0, bottom: 0 } };
    const bodyPr = txBody.getElementsByTagNameNS("*", "bodyPr")[0] ?? null;
    let vertical = "top";
    const tcPr = tc?.getElementsByTagNameNS("*", "tcPr")[0] ?? null;
    const tcAnchor = tcPr?.getAttribute("anchor") || tcPr?.getAttribute("vAlign") || void 0;
    const bpAnchor = bodyPr?.getAttribute("anchor") || void 0;
    const vSrc = tcAnchor || bpAnchor;
    if (vSrc === "ctr") vertical = "middle";
    else if (vSrc === "b") vertical = "bottom";
    else vertical = "top";
    const lIns = bodyPr?.getAttribute("lIns");
    const tIns = bodyPr?.getAttribute("tIns");
    const rIns = bodyPr?.getAttribute("rIns");
    const bIns = bodyPr?.getAttribute("bIns");
    const padding = {
      left: lIns ? Number(lIns) / 9525 : 6,
      top: tIns ? Number(tIns) / 9525 : 2,
      right: rIns ? Number(rIns) / 9525 : 6,
      bottom: bIns ? Number(bIns) / 9525 : 2
    };
    let horiz = "left";
    let fontName = "Arial";
    let fontSize = 14;
    let color;
    let parts = [];
    const paragraphs = txBody.getElementsByTagNameNS("*", "p");
    for (const p of Array.from(paragraphs)) {
      const pPr = p.getElementsByTagNameNS("*", "pPr")[0] ?? null;
      const algn = pPr?.getAttribute("algn") || void 0;
      if (algn) {
        horiz = algn === "ctr" ? "center" : algn === "r" ? "right" : algn.startsWith("just") ? "justify" : "left";
      }
      for (const child of Array.from(p.childNodes)) {
        if (child.nodeType === 1) {
          const ln = child.localName;
          if (ln === "r") {
            const rPr = child.getElementsByTagNameNS("*", "rPr")[0] ?? null;
            if (rPr) {
              const latin = rPr.getElementsByTagNameNS("*", "latin")[0] ?? null;
              fontName = latin?.getAttribute("typeface") ?? fontName;
              const sz = rPr.getAttribute("sz");
              if (sz) fontSize = parseInt(sz, 10) / 100;
              const solidFill = rPr.querySelector("*|solidFill");
              const c = XmlHelper.getColorFromElement(solidFill || null, themeColors);
              if (c) color = c;
            }
            const t = child.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
            parts.push(t);
          } else if (ln === "br") {
            parts.push("\n");
          } else if (ln === "fld") {
            const runs = child.getElementsByTagNameNS("*", "r");
            for (const r of Array.from(runs)) {
              const t = r.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
              parts.push(t);
            }
          }
        }
      }
    }
    return {
      text: parts.join("").trim(),
      font: { name: fontName, size: fontSize, color },
      align: { horizontal: horiz, vertical },
      padding
    };
  }
  static extractFillColor(tcPr, themeColors) {
    if (!tcPr) return void 0;
    if (tcPr.getElementsByTagNameNS("*", "noFill")[0]) return void 0;
    const direct = Array.from(tcPr.children).find((c) => c.localName === "solidFill");
    if (direct) return XmlHelper.getColorFromElement(direct, themeColors);
    const all = Array.from(tcPr.getElementsByTagNameNS("*", "solidFill"));
    for (const cand of all) {
      let p = cand.parentElement;
      let insideLine = false;
      while (p && p !== tcPr) {
        const ln = p.localName;
        if (ln === "ln" || ln === "lnL" || ln === "lnR" || ln === "lnT" || ln === "lnB" || ln === "tcBorders") {
          insideLine = true;
          break;
        }
        p = p.parentElement;
      }
      if (!insideLine) {
        const col = XmlHelper.getColorFromElement(cand, themeColors);
        if (col) return col;
      }
    }
    return void 0;
  }
  static extractCellBorders(tcPr, themeColors) {
    const borders = {};
    if (!tcPr) return borders;
    const map = {
      lnL: "left",
      lnR: "right",
      lnT: "top",
      lnB: "bottom"
    };
    for (const key of Object.keys(map)) {
      const ln = tcPr.getElementsByTagNameNS("*", key)[0] ?? null;
      if (!ln) continue;
      const wAttr = ln.getAttribute("w");
      const w = wAttr ? Number(wAttr) / 9525 : void 0;
      const solidFill = ln.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
      const color = XmlHelper.getColorFromElement(solidFill, themeColors);
      const prstDash = ln.getElementsByTagNameNS("*", "prstDash")[0] ?? null;
      const dashVal = prstDash?.getAttribute("val") || "";
      const style = this.mapPrstDashToCss(dashVal);
      borders[map[key]] = { color, width: w, style };
    }
    return borders;
  }
  static extractTableStyleFlags(tblPr) {
    if (!tblPr) return {};
    const flags = {};
    for (const k of ["firstRow", "firstCol", "lastRow", "lastCol", "bandRow", "bandCol"]) {
      if (tblPr.getAttribute(k) === "1" || tblPr.getAttribute(k) === "true") flags[k] = true;
    }
    return flags;
  }
  static extractTableStyleId(tblPr) {
    if (!tblPr) return void 0;
    const idEl = tblPr.getElementsByTagNameNS("*", "tblStyleId")[0] || tblPr.getElementsByTagNameNS("*", "tblStyle")[0] || null;
    const idAttr = idEl?.getAttribute("val") || void 0;
    const text = idEl?.textContent?.trim() || void 0;
    const direct = tblPr.getAttribute("tblStyle") || void 0;
    return idAttr || text || direct || void 0;
  }
  static extractTableBorders(tblPr, themeColors) {
    if (!tblPr) return void 0;
    const borders = {};
    const tblBorders = tblPr.getElementsByTagNameNS("*", "tblBorders")[0] ?? null;
    if (!tblBorders) return void 0;
    const map = {
      top: "top",
      right: "right",
      bottom: "bottom",
      left: "left",
      insideH: "insideH",
      insideV: "insideV"
    };
    for (const tag of Object.keys(map)) {
      const node = tblBorders.getElementsByTagNameNS("*", tag)[0] ?? null;
      if (!node) continue;
      const ln = node.getElementsByTagNameNS("*", "ln")[0] ?? node;
      const wAttr = ln.getAttribute("w");
      const w = wAttr ? Number(wAttr) / 9525 : void 0;
      const solidFill = ln.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
      const color = XmlHelper.getColorFromElement(solidFill, themeColors);
      const prstDash = ln.getElementsByTagNameNS("*", "prstDash")[0] ?? null;
      const dashVal = prstDash?.getAttribute("val") || "";
      const style = this.mapPrstDashToCss(dashVal);
      borders[map[tag]] = { color, width: w, style };
    }
    return borders;
  }
  static extractTableFill(tblPr, themeColors) {
    if (!tblPr) return void 0;
    const directSolid = Array.from(tblPr.children).find((c) => c.localName === "solidFill");
    if (directSolid) return XmlHelper.getColorFromElement(directSolid, themeColors);
    const allSolid = Array.from(tblPr.getElementsByTagNameNS("*", "solidFill"));
    for (const cand of allSolid) {
      let p = cand.parentElement;
      let insideBorders = false;
      while (p && p !== tblPr) {
        if (p.localName === "tblBorders") {
          insideBorders = true;
          break;
        }
        p = p.parentElement;
      }
      if (!insideBorders) {
        const col = XmlHelper.getColorFromElement(cand, themeColors);
        if (col) return col;
      }
    }
    return void 0;
  }
  static mapPrstDashToCss(val) {
    const v = (val || "").toLowerCase();
    if (!v) return void 0;
    if (v === "solid") return "solid";
    if (v.includes("dot")) return "dotted";
    if (v.includes("dash")) return "dashed";
    return void 0;
  }
};

// src/elements/ChartExtractor.ts
var ChartExtractor = class {
  static async extract(spTree, relsXml, zip, themeColors) {
    if (!spTree) return [];
    const charts = [];
    const gFrames = spTree.getElementsByTagNameNS("*", "graphicFrame");
    for (const gf of Array.from(gFrames)) {
      const graphicData = gf.getElementsByTagNameNS("*", "graphicData")[0] ?? null;
      if (!graphicData) continue;
      const chartEl = graphicData.getElementsByTagNameNS("*", "chart")[0] ?? null;
      if (!chartEl) continue;
      const rId = chartEl.getAttribute("r:id") || chartEl.getAttribute("r:embed") || void 0;
      if (!rId) continue;
      const rel = XmlHelper.findRelationshipById(relsXml, rId);
      const target = rel?.getAttribute("Target") || void 0;
      if (!target) continue;
      const fullPath = this.resolvePath(target, "ppt/slides");
      const file = zip.file(fullPath);
      if (!file) continue;
      const xmlStr = await file.async("string");
      const doc = XmlHelper.parseXml(xmlStr);
      const parsed = this.parseChart(doc, themeColors);
      if (!parsed) continue;
      const xfrm = gf.getElementsByTagNameNS("*", "xfrm")[0] ?? null;
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;
      const x = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
      const y = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;
      const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1e6;
      const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 6e5;
      charts.push({
        type: "chart",
        chartType: parsed.type,
        position: { x, y },
        size: { width: cx, height: cy },
        categories: parsed.categories,
        series: parsed.series,
        palette: parsed.palette,
        title: parsed.title,
        showLegend: parsed.showLegend,
        showDataLabels: parsed.showDataLabels,
        stackedMode: parsed.stackedMode,
        valueFormat: parsed.valueFormat
      });
    }
    return charts;
  }
  static resolvePath(target, baseDir) {
    const parts = (baseDir + "/" + target).split("/");
    const resolved = [];
    for (const part of parts) {
      if (part === "..") resolved.pop();
      else if (part !== ".") resolved.push(part);
    }
    return resolved.join("/");
  }
  static parseChart(doc, themeColors) {
    const plotArea = doc.getElementsByTagNameNS("*", "plotArea")[0] || null;
    if (!plotArea) return null;
    const titleText = this.extractTitle(doc);
    const showLegend = !!doc.getElementsByTagNameNS("*", "legend")[0];
    const showDataLabels = !!plotArea.getElementsByTagNameNS("*", "dLbls")[0];
    const bar = plotArea.getElementsByTagNameNS("*", "barChart")[0] || null;
    const line = plotArea.getElementsByTagNameNS("*", "lineChart")[0] || null;
    const area = plotArea.getElementsByTagNameNS("*", "areaChart")[0] || null;
    const pie = plotArea.getElementsByTagNameNS("*", "pieChart")[0] || null;
    const scatter = plotArea.getElementsByTagNameNS("*", "scatterChart")[0] || null;
    const chartNumFmt = plotArea.getElementsByTagNameNS("*", "dLbls")[0]?.getElementsByTagNameNS("*", "numFmt")[0]?.getAttribute("formatCode") || void 0;
    const palette = [
      themeColors["accent1"],
      themeColors["accent2"],
      themeColors["accent3"],
      themeColors["accent4"],
      themeColors["accent5"],
      themeColors["accent6"]
    ].filter(Boolean);
    if (bar) {
      const cat = this.extractCategories(bar) || [];
      const ser = this.extractSeries(bar, themeColors) || [];
      const barDir = bar.getElementsByTagNameNS("*", "barDir")[0]?.getAttribute("val") || "col";
      const type = barDir === "bar" ? "bar" : "column";
      const grouping = bar.getElementsByTagNameNS("*", "grouping")[0]?.getAttribute("val") || "clustered";
      const stackedMode = grouping === "stacked" ? "stacked" : grouping === "percentStacked" ? "percent" : "none";
      return { type, categories: cat, series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode, valueFormat: chartNumFmt };
    }
    if (line) {
      const cat = this.extractCategories(line) || [];
      const ser = this.extractSeries(line, themeColors) || [];
      const grouping = line.getElementsByTagNameNS("*", "grouping")[0]?.getAttribute("val") || "standard";
      const stackedMode = grouping === "stacked" ? "stacked" : grouping === "percentStacked" ? "percent" : "none";
      return { type: "line", categories: cat, series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode, valueFormat: chartNumFmt };
    }
    if (area) {
      const cat = this.extractCategories(area) || [];
      const ser = this.extractSeries(area, themeColors) || [];
      const grouping = area.getElementsByTagNameNS("*", "grouping")[0]?.getAttribute("val") || "standard";
      const stackedMode = grouping === "stacked" ? "stacked" : grouping === "percentStacked" ? "percent" : "none";
      return { type: "area", categories: cat, series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode, valueFormat: chartNumFmt };
    }
    if (pie) {
      const cat = this.extractCategories(pie) || [];
      const ser = this.extractSeries(pie, themeColors) || [];
      return { type: "pie", categories: cat, series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode: "none", valueFormat: chartNumFmt };
    }
    if (scatter) {
      const ser = this.extractScatterSeries(scatter, themeColors) || [];
      return { type: "scatter", categories: [], series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode: "none", valueFormat: chartNumFmt };
    }
    return null;
  }
  static extractTitle(doc) {
    const title = doc.getElementsByTagNameNS("*", "title")[0] || null;
    if (!title) return void 0;
    const tx = title.getElementsByTagNameNS("*", "tx")[0] || null;
    const rich = tx?.getElementsByTagNameNS("*", "rich")[0] || null;
    if (rich) {
      const t = rich.getElementsByTagNameNS("*", "t")[0]?.textContent || void 0;
      return t || void 0;
    }
    const v = tx?.getElementsByTagNameNS("*", "v")[0]?.textContent || void 0;
    return v || void 0;
  }
  static extractCategories(parent) {
    const cat = parent.getElementsByTagNameNS("*", "cat")[0] || null;
    if (!cat) return null;
    const strCache = cat.getElementsByTagNameNS("*", "strCache")[0] || null;
    if (strCache) {
      const pts = Array.from(strCache.getElementsByTagNameNS("*", "pt"));
      return pts.map((p) => p.getElementsByTagNameNS("*", "v")[0]?.textContent || "");
    }
    const numCache = cat.getElementsByTagNameNS("*", "numCache")[0] || null;
    if (numCache) {
      const pts = Array.from(numCache.getElementsByTagNameNS("*", "pt"));
      return pts.map((p) => Number(p.getElementsByTagNameNS("*", "v")[0]?.textContent || 0));
    }
    return null;
  }
  static extractSeries(parent, themeColors) {
    const series = [];
    const sers = Array.from(parent.getElementsByTagNameNS("*", "ser"));
    let idx = 0;
    for (const s of sers) {
      const name = s.getElementsByTagNameNS("*", "tx")[0]?.getElementsByTagNameNS("*", "v")[0]?.textContent || void 0;
      const numCache = s.getElementsByTagNameNS("*", "numCache")[0] || null;
      let values = [];
      if (numCache) {
        const pts = Array.from(numCache.getElementsByTagNameNS("*", "pt"));
        values = pts.map((p) => Number(p.getElementsByTagNameNS("*", "v")[0]?.textContent || 0));
      }
      const valueFormat = s.getElementsByTagNameNS("*", "dLbls")[0]?.getElementsByTagNameNS("*", "numFmt")[0]?.getAttribute("formatCode") || void 0;
      const spPr = s.getElementsByTagNameNS("*", "spPr")[0] || null;
      const solidFill = spPr?.getElementsByTagNameNS("*", "solidFill")[0] || null;
      const color = XmlHelper.getColorFromElement(solidFill, themeColors);
      series.push({ name, values, color, valueFormat });
      idx += 1;
    }
    return series;
  }
  static extractScatterSeries(parent, themeColors) {
    const out = [];
    const sers = Array.from(parent.getElementsByTagNameNS("*", "ser"));
    for (const s of sers) {
      const name = s.getElementsByTagNameNS("*", "tx")[0]?.getElementsByTagNameNS("*", "v")[0]?.textContent || void 0;
      const xCache = s.getElementsByTagNameNS("*", "xVal")[0]?.getElementsByTagNameNS("*", "numCache")[0] || null;
      const yCache = s.getElementsByTagNameNS("*", "yVal")[0]?.getElementsByTagNameNS("*", "numCache")[0] || null;
      const xPts = xCache ? Array.from(xCache.getElementsByTagNameNS("*", "pt")) : [];
      const yPts = yCache ? Array.from(yCache.getElementsByTagNameNS("*", "pt")) : [];
      const len = Math.min(xPts.length, yPts.length);
      const points = [];
      for (let i = 0; i < len; i++) {
        const xv = Number(xPts[i].getElementsByTagNameNS("*", "v")[0]?.textContent || 0);
        const yv = Number(yPts[i].getElementsByTagNameNS("*", "v")[0]?.textContent || 0);
        points.push({ x: xv, y: yv });
      }
      const spPr = s.getElementsByTagNameNS("*", "spPr")[0] || null;
      const solidFill = spPr?.getElementsByTagNameNS("*", "solidFill")[0] || null;
      const color = XmlHelper.getColorFromElement(solidFill, themeColors);
      const valueFormat = s.getElementsByTagNameNS("*", "dLbls")[0]?.getElementsByTagNameNS("*", "numFmt")[0]?.getAttribute("formatCode") || void 0;
      out.push({ name, points, color, valueFormat });
    }
    return out;
  }
};

// src/core/SlideExtractor.ts
var SlideExtractor = class {
  constructor(zip) {
    this.zip = zip;
  }
  /**
   * Extracts all slides in order and parses their visual elements.
   * @returns An array of SlideElement lists (one per slide).
   */
  async extractSlides() {
    const themeXmlStr = await this.zip.file("ppt/theme/theme1.xml")?.async("string");
    const themeXml = themeXmlStr ? XmlHelper.parseXml(themeXmlStr) : null;
    const themeColors = XmlHelper.extractThemeColors(themeXml);
    const themeTableStyles = XmlHelper.extractThemeTableStyles(themeXml);
    const slidePaths = Object.keys(this.zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] || "0", 10);
      const numB = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] || "0", 10);
      return numA - numB;
    });
    const slides = [];
    for (const slidePath of slidePaths) {
      const slideXmlStr = await this.zip.file(slidePath).async("string");
      const slideXml = XmlHelper.parseXml(slideXmlStr);
      const spTree = slideXml.querySelector("p\\:spTree") || slideXml.getElementsByTagNameNS("*", "spTree")[0];
      if (!spTree) {
        console.warn(`Warning: no <spTree> found in ${slidePath}`);
        slides.push([]);
        continue;
      }
      const relsPath = slidePath.replace("slides/", "slides/_rels/") + ".rels";
      const relsXml = this.zip.file(relsPath) ? XmlHelper.parseXml(await this.zip.file(relsPath).async("string")) : XmlHelper.parseXml(`<Relationships/>`);
      const layoutRel = XmlHelper.findRelationshipByTypeSuffix(relsXml, "/slideLayout");
      const layoutTarget = layoutRel?.getAttribute("Target") || void 0;
      let layoutSpTree = null;
      let layoutRelsXml = null;
      if (layoutTarget) {
        const layoutPath = this.resolvePath(layoutTarget, "ppt/slides");
        const layoutXmlStr = await this.zip.file(layoutPath)?.async("string");
        if (layoutXmlStr) {
          const layoutXml = XmlHelper.parseXml(layoutXmlStr);
          layoutSpTree = layoutXml.querySelector("p\\:spTree") || layoutXml.getElementsByTagNameNS("*", "spTree")[0] || null;
          const layoutRelsPath = layoutPath.replace("slideLayouts/", "slideLayouts/_rels/") + ".rels";
          layoutRelsXml = this.zip.file(layoutRelsPath) ? XmlHelper.parseXml(await this.zip.file(layoutRelsPath).async("string")) : XmlHelper.parseXml(`<Relationships/>`);
        }
      }
      let masterSpTree = null;
      let masterRelsXml = null;
      if (layoutRelsXml) {
        const masterRel = XmlHelper.findRelationshipByTypeSuffix(layoutRelsXml, "/slideMaster");
        const masterTarget = masterRel?.getAttribute("Target") || void 0;
        if (masterTarget) {
          const masterPath = this.resolvePath(masterTarget, "ppt/slideLayouts");
          const masterXmlStr = await this.zip.file(masterPath)?.async("string");
          if (masterXmlStr) {
            const masterXml = XmlHelper.parseXml(masterXmlStr);
            masterSpTree = masterXml.querySelector("p\\:spTree") || masterXml.getElementsByTagNameNS("*", "spTree")[0] || null;
            const masterRelsPath = masterPath.replace("slideMasters/", "slideMasters/_rels/") + ".rels";
            masterRelsXml = this.zip.file(masterRelsPath) ? XmlHelper.parseXml(await this.zip.file(masterRelsPath).async("string")) : XmlHelper.parseXml(`<Relationships/>`);
          }
        }
      }
      const slideBg = await this.extractBackground(slideXml, relsXml, "ppt/slides", this.zip, themeColors);
      const layoutBg = layoutRelsXml ? await this.extractBackground(layoutSpTree?.ownerDocument || null, layoutRelsXml, "ppt/slideLayouts", this.zip, themeColors) : null;
      const masterBg = masterRelsXml ? await this.extractBackground(masterSpTree?.ownerDocument || null, masterRelsXml, "ppt/slideMasters", this.zip, themeColors) : null;
      const bgElement = slideBg || layoutBg || masterBg;
      const masterText = masterSpTree ? TextExtractor.extract(masterSpTree, themeColors, { context: "master" }) : [];
      const masterImages = masterSpTree && masterRelsXml ? await ImageExtractor.extract(masterSpTree, masterRelsXml, this.zip, "ppt/slideMasters") : [];
      const masterShapes = masterSpTree ? ShapeExtractor.extract(masterSpTree, themeColors) : [];
      const layoutText = layoutSpTree ? TextExtractor.extract(layoutSpTree, themeColors, { context: "layout" }) : [];
      const layoutImages = layoutSpTree && layoutRelsXml ? await ImageExtractor.extract(layoutSpTree, layoutRelsXml, this.zip, "ppt/slideLayouts") : [];
      const layoutShapes = layoutSpTree ? ShapeExtractor.extract(layoutSpTree, themeColors) : [];
      const masterGeom = this.extractPlaceholderGeom(masterSpTree);
      const layoutGeom = this.extractPlaceholderGeom(layoutSpTree);
      const mergedGeom = { ...masterGeom, ...layoutGeom };
      const slideText = TextExtractor.extract(spTree, themeColors, { context: "slide", placeholderGeom: mergedGeom });
      const slideImages = await ImageExtractor.extract(spTree, relsXml, this.zip, "ppt/slides");
      const slideTables = TableExtractor.extract(spTree, themeColors, themeTableStyles);
      const slideCharts = await ChartExtractor.extract(spTree, relsXml, this.zip, themeColors);
      const slideShapes = ShapeExtractor.extract(spTree, themeColors);
      slides.push([
        ...bgElement ? [bgElement] : [],
        ...masterShapes,
        ...masterImages,
        ...masterText,
        ...layoutShapes,
        ...layoutImages,
        ...layoutText,
        ...slideShapes,
        ...slideTables,
        ...slideCharts,
        ...slideImages,
        ...slideText
      ]);
    }
    return slides;
  }
  /** Normalize a relative path against a base directory inside ppt folder */
  resolvePath(target, baseDir) {
    const parts = (baseDir + "/" + target).split("/");
    const resolved = [];
    for (const part of parts) {
      if (part === "..") {
        if (resolved.length) resolved.pop();
      } else if (part !== "." && part !== "") {
        resolved.push(part);
      }
    }
    return resolved.join("/");
  }
  extractPlaceholderGeom(spTree) {
    const map = {};
    if (!spTree) return map;
    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const nvPr = shape.getElementsByTagNameNS("*", "nvPr")[0] ?? null;
      const ph = nvPr?.getElementsByTagNameNS("*", "ph")[0] ?? null;
      const idx = ph?.getAttribute("idx") || void 0;
      const phType = ph?.getAttribute("type") || void 0;
      const key = idx || phType;
      if (!key) continue;
      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0] ?? null;
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;
      if (!off || !ext) continue;
      let fontSize = void 0;
      const txBody = shape.getElementsByTagNameNS("*", "txBody")[0] ?? null;
      const defRPr = txBody?.querySelector("*|lstStyle *|defRPr");
      const sz = defRPr?.getAttribute("sz");
      if (sz) {
        const n = parseInt(sz, 10);
        if (Number.isFinite(n)) fontSize = n / 100;
      }
      map[key] = {
        x: XmlHelper.getAttrAsNumber(off, "x"),
        y: XmlHelper.getAttrAsNumber(off, "y"),
        cx: XmlHelper.getAttrAsNumber(ext, "cx"),
        cy: XmlHelper.getAttrAsNumber(ext, "cy"),
        fontSize
      };
    }
    return map;
  }
  async extractBackground(doc, rels, baseDir, zip, themeColors) {
    if (!doc) return null;
    const bg = doc.getElementsByTagNameNS("*", "bg")[0];
    if (!bg) return null;
    const bgPr = bg.getElementsByTagNameNS("*", "bgPr")[0] || null;
    const solidFill = bgPr?.getElementsByTagNameNS("*", "solidFill")[0] || null;
    const color = XmlHelper.getColorFromElement(solidFill, themeColors);
    if (color) {
      return { type: "background", fill: { type: "solid", color } };
    }
    const gradFill = bgPr?.getElementsByTagNameNS("*", "gradFill")[0] || null;
    const gradient = XmlHelper.getGradientFromElement(gradFill, themeColors);
    if (gradient) {
      return { type: "background", fill: gradient };
    }
    const bgRef = bg.getElementsByTagNameNS("*", "bgRef")[0] || null;
    const schemeClr = bgRef?.getElementsByTagNameNS("*", "schemeClr")[0] || null;
    const schemeVal = schemeClr?.getAttribute("val") || void 0;
    if (schemeVal && themeColors[schemeVal]) {
      return { type: "background", fill: { type: "solid", color: themeColors[schemeVal] } };
    }
    const blipFill = bgPr?.getElementsByTagNameNS("*", "blipFill")[0] || null;
    const blip = blipFill?.getElementsByTagNameNS("*", "blip")[0] || null;
    const embedId = blip?.getAttribute("r:embed") || void 0;
    if (embedId && rels) {
      const rel = XmlHelper.findRelationshipById(rels, embedId);
      const target = rel?.getAttribute("Target") || void 0;
      if (target) {
        const fullPath = this.resolvePath(target, baseDir);
        const file = zip.file(fullPath);
        if (file) {
          const binary = await file.async("base64");
          const ext = fullPath.split(".").pop()?.toLowerCase() || "png";
          const dataUri = `data:image/${ext};base64,${binary}`;
          return { type: "background", fill: { type: "image", src: dataUri } };
        }
      }
    }
    return null;
  }
};

// src/core/PptxReader.ts
var PptxReader = class {
  zip;
  baseWidthPx;
  baseHeightPx;
  /**
   * Loads and parses a .pptx binary buffer.
   * @param buffer The binary content of a .pptx file.
   * @returns A list of slides, each represented as an array of SlideElement.
   */
  async load(buffer) {
    this.zip = await JSZip.loadAsync(buffer);
    await this.computeSlideBaseSize();
    const extractor = new SlideExtractor(this.zip);
    return extractor.extractSlides();
  }
  /**
   * Returns slide base size in pixels derived from ppt/presentation.xml sldSz (if available).
   * Defaults to 960x540 when not found.
   */
  async getBaseSizePx() {
    if (this.baseWidthPx && this.baseHeightPx) {
      return { width: this.baseWidthPx, height: this.baseHeightPx };
    }
    await this.computeSlideBaseSize();
    return {
      width: Number.isFinite(this.baseWidthPx) && this.baseWidthPx > 0 ? this.baseWidthPx : 960,
      height: Number.isFinite(this.baseHeightPx) && this.baseHeightPx > 0 ? this.baseHeightPx : 540
    };
  }
  async computeSlideBaseSize() {
    try {
      if (!this.zip) return;
      const presFile = this.zip.file("ppt/presentation.xml");
      if (!presFile) return;
      const xmlStr = await presFile.async("string");
      const doc = XmlHelper.parseXml(xmlStr);
      const sldSz = doc.getElementsByTagNameNS("*", "sldSz")[0];
      const cx = sldSz ? Number(sldSz.getAttribute("cx") || 0) : 0;
      const cy = sldSz ? Number(sldSz.getAttribute("cy") || 0) : 0;
      if (Number.isFinite(cx) && Number.isFinite(cy) && cx > 0 && cy > 0) {
        this.baseWidthPx = cx / 9525;
        this.baseHeightPx = cy / 9525;
      }
    } catch {
    }
  }
};

// src/renderer/renderTextElement.ts
function renderTextElement(el) {
  const nf = (n, fb = 0) => Number.isFinite(n) ? n : fb;
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const w = nf(el.size?.width, 0) / 9525;
  const h = nf(el.size?.height, 0) / 9525;
  const pad = el.padding || { left: 0, top: 0, right: 0, bottom: 0 };
  const textAlign = el.align?.horizontal || "left";
  const justify = el.align?.vertical === "middle" ? "center" : el.align?.vertical === "bottom" ? "flex-end" : "flex-start";
  const inner = el.html ? el.html : escape(el.content);
  return `<div style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${w}px;
    height: ${h}px;
    display: flex;
    flex-direction: column;
    justify-content: ${justify};
    text-align: ${textAlign};
    padding: ${pad.top}px ${pad.right}px ${pad.bottom}px ${pad.left}px;
    font-family: ${el.font?.name || "Arial"}, sans-serif;
    font-size: ${nf(Number(el.font?.size), 12)}pt;
    color: ${el.font?.color || "#000"};
    overflow: visible;
    white-space: pre-wrap;
  ">${inner}</div>`;
}
function escape(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// src/renderer/renderImageElement.ts
function renderImageElement(el) {
  const nf = (n, fb = 0) => Number.isFinite(n) ? n : fb;
  return `<img src="${el.src}" style="
    position: absolute;
    left: ${nf(el.position?.x, 0) / 9525}px;
    top: ${nf(el.position?.y, 0) / 9525}px;
    width: ${nf(el.size?.width, 0) / 9525}px;
    height: ${nf(el.size?.height, 0) / 9525}px;
    object-fit: cover;
  " />`;
}

// src/renderer/shapePathMap.ts
function getSvgPathForShape(type) {
  switch (type) {
    // ▸ Basic Arrows
    case "rightArrow":
    case "arrow":
      return "POLYGON 0,25 70,25 70,0 100,50 70,100 70,75 0,75";
    case "leftArrow":
      return "POLYGON 100,25 30,25 30,0 0,50 30,100 30,75 100,75";
    case "leftRightArrow":
      return "POLYGON 0,50 30,0 30,25 70,25 70,0 100,50 70,100 70,75 30,75 30,100";
    case "triangle":
      return "POLYGON 50,0 100,100 0,100";
    case "star5":
      return "POLYGON 50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35";
    case "cloud":
      return "PATH M20,60 C10,60 10,40 25,40 C30,20 50,20 55,35 C70,30 80,40 80,50 C90,50 90,70 75,70 H25 Z";
    // ▸ Straight connector / line
    case "line":
    case "straightConnector1":
      return "LINE_ARROW 0,0 100,100";
    // ▸ Bent connectors
    case "bentConnector2":
      return "POLYLINE 0,50 50,50 50,100";
    case "bentConnector3":
      return "POLYLINE_ARROW 0,50 40,50 40,70 100,70";
    case "bentConnector4":
      return "POLYLINE 0,30 30,30 30,70 70,70 70,100";
    case "bentConnector5":
      return "POLYLINE 0,20 30,20 30,50 60,50 60,80 100,80";
    // ▸ Curved connectors
    case "curvedConnector2":
      return "PATH M0,50 Q50,0 100,50";
    case "curvedConnector3":
      return "PATH M0,50 Q25,0 50,50 Q75,100 100,50";
    case "curvedConnector4":
      return "PATH M0,40 Q20,0 40,40 Q60,80 80,40 Q90,20 100,40";
    case "curvedConnector5":
      return "PATH M0,50 Q20,20 40,50 Q60,80 80,50 Q90,40 100,50";
    // ▸ Notched, bent, and curved arrows
    case "bentArrow":
      return "POLYGON 0,0 70,0 70,30 100,30 50,100 50,30 0,30";
    case "notchedRightArrow":
      return "POLYGON 0,20 60,20 60,0 100,50 60,100 60,80 0,80";
    case "curvedRightArrow":
      return "PATH M0,50 Q50,0 100,50 Q50,100 0,50 Z";
    // Default fallback (rectangular shape)
    default:
      return "POLYGON 0,0 100,0 100,100 0,100";
  }
}

// src/renderer/renderShapeElement.ts
function fillToCssColor(fill, fallback = "transparent") {
  switch (fill.type) {
    case "solid":
      return fill.color;
    case "gradient":
      return fill.stops.length > 0 ? fill.stops[0].color : fallback;
    case "none":
      return "transparent";
    default:
      return fallback;
  }
}
function fillToCssBackground(fill) {
  switch (fill.type) {
    case "solid":
      return `background-color: ${fill.color};`;
    case "gradient": {
      if (fill.stops.length === 0) return "background-color: transparent;";
      if (fill.stops.length === 1) return `background-color: ${fill.stops[0].color};`;
      const stops = fill.stops.map((s) => `${s.color} ${s.offset}%`).join(", ");
      if (fill.gradientType === "radial") {
        return `background: radial-gradient(ellipse at center, ${stops});`;
      }
      const angle = fill.angle ?? 180;
      return `background: linear-gradient(${angle}deg, ${stops});`;
    }
    case "image":
      return `background-image: url('${fill.src}'); background-size: cover; background-position: center;`;
    case "none":
      return "background-color: transparent;";
  }
}
function dashStyleToSvg(dashStyle) {
  if (!dashStyle || dashStyle === "solid") return void 0;
  const map = {
    dash: "8,4",
    dot: "2,4",
    dashDot: "8,4,2,4",
    lgDash: "16,4",
    lgDashDot: "16,4,2,4",
    lgDashDotDot: "16,4,2,4,2,4",
    sysDash: "4,2",
    sysDot: "1,2",
    sysDashDot: "4,2,1,2",
    sysDashDotDot: "4,2,1,2,1,2"
  };
  return map[dashStyle];
}
function resolveStrokeColor(el) {
  const strokeColor = el.stroke?.color;
  if (strokeColor && strokeColor !== "transparent") return strokeColor;
  const fc = fillToCssColor(el.fill);
  if (fc !== "transparent") return fc;
  return "#000";
}
function renderShapeElement(el, options = {}) {
  const nf = (n, fb = 0) => Number.isFinite(n) ? n : fb;
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const width = nf(el.size?.width, 0) / 9525;
  const height = nf(el.size?.height, 0) / 9525;
  const rotation = el.rotationDeg && !isNaN(el.rotationDeg) ? el.rotationDeg : 0;
  const rotationStyle = rotation ? `transform: rotate(${rotation}deg); transform-origin: center;` : "";
  const strokeWidth = el.stroke?.width && Number.isFinite(el.stroke.width) && el.stroke.width > 0 ? el.stroke.width : 1;
  const strokeColor = el.stroke?.color ?? "transparent";
  const style = `
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${width}px;
    height: ${height}px;
    ${rotationStyle}
  `;
  const bgCss = fillToCssBackground(el.fill);
  if (el.shapeType === "rect") {
    return `<div style="${style}
      ${bgCss}
      border: ${strokeWidth}px solid ${strokeColor};
      box-sizing: border-box;"></div>`;
  }
  if (el.shapeType === "ellipse") {
    return `<div style="${style}
      ${bgCss}
      border: ${strokeWidth}px solid ${strokeColor};
      border-radius: 50%;
      box-sizing: border-box;"></div>`;
  }
  if (el.shapeType === "line") {
    const sw = el.stroke?.width && Number.isFinite(el.stroke.width) && el.stroke.width > 0 ? el.stroke.width : 1;
    const lineColor = resolveStrokeColor(el);
    const dashAttr = dashStyleToSvg(el.stroke?.dashStyle);
    const dashStr = dashAttr ? ` stroke-dasharray="${dashAttr}"` : "";
    const defs = buildMarkerDefs(el.stroke?.headEnd, el.stroke?.tailEnd, lineColor);
    const markerStartAttr = defs.startId ? ` marker-start="url(#${defs.startId})"` : "";
    const markerEndAttr = defs.endId ? ` marker-end="url(#${defs.endId})"` : "";
    const startPull = defs.startLen * sw;
    const endPull = defs.endLen * sw;
    const svgW = Math.max(width, sw * 2);
    const svgH = Math.max(height, sw * 2);
    const isVertical = width < 1;
    const isHorizontal = height < 1;
    const fh = el.flipH ?? false;
    const fv = el.flipV ?? false;
    let x1, y1, x2, y2;
    if (isVertical) {
      const cx = svgW / 2;
      const topY = startPull;
      const botY = svgH - endPull;
      if (fv) {
        x1 = cx;
        y1 = svgH - startPull;
        x2 = cx;
        y2 = endPull;
      } else {
        x1 = cx;
        y1 = topY;
        x2 = cx;
        y2 = botY;
      }
    } else if (isHorizontal) {
      const cy = svgH / 2;
      if (fh) {
        x1 = svgW - startPull;
        y1 = cy;
        x2 = endPull;
        y2 = cy;
      } else {
        x1 = startPull;
        y1 = cy;
        x2 = svgW - endPull;
        y2 = cy;
      }
    } else {
      let sx = 0, sy = 0, ex = svgW, ey = svgH;
      if (fh) {
        sx = svgW;
        ex = 0;
      }
      if (fv) {
        sy = svgH;
        ey = 0;
      }
      const dx = ex - sx, dy = ey - sy;
      const lineLen = Math.sqrt(dx * dx + dy * dy);
      const ux = dx / lineLen, uy = dy / lineLen;
      x1 = sx + startPull * ux;
      y1 = sy + startPull * uy;
      x2 = ex - endPull * ux;
      y2 = ey - endPull * uy;
    }
    return `<svg viewBox="0 0 ${svgW} ${svgH}" style="
              position: absolute;
              left: ${isVertical ? x - svgW / 2 : x}px;
              top: ${isHorizontal ? y - svgH / 2 : y}px;
              width: ${svgW}px;
              height: ${svgH}px;
              ${rotationStyle}
            " overflow="visible">
          ${defs.defs}
          <line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}"
                stroke="${lineColor}" stroke-width="${sw}"${dashStr}${markerStartAttr}${markerEndAttr} />
        </svg>`;
  }
  if (el.shapeType === "roundRect") {
    const adjVal = el.cornerRadiusPct;
    let radius;
    if (adjVal !== void 0 && adjVal >= 0) {
      const shorter = Math.min(width, height);
      radius = adjVal / 1e5 * shorter;
    } else {
      radius = Math.min(16, Math.min(width, height) * 0.1);
    }
    return `<div style="${style}
      ${bgCss}
      border: ${strokeWidth}px solid ${strokeColor};
      border-radius: ${radius}px;
      box-sizing: border-box;"></div>`;
  }
  const raw = getSvgPathForShape(el.shapeType);
  const svgFill = fillToCssColor(el.fill);
  const svgStroke = resolveStrokeColor(el);
  return shapeSvg(
    x,
    y,
    width,
    height,
    svgFill,
    svgStroke,
    raw,
    el.stroke?.width && Number.isFinite(el.stroke.width) ? el.stroke.width : void 0,
    rotation,
    el.stroke?.headEnd,
    el.stroke?.tailEnd,
    options.scaleStrokes === true,
    dashStyleToSvg(el.stroke?.dashStyle)
  );
}
function shapeSvg(x, y, width, height, fill, stroke, raw, strokeWidthPx, rotationDeg, headEnd, tailEnd, scaleStrokes, dashArray) {
  const strokeColorOpt = stroke && stroke !== "transparent" ? stroke : void 0;
  const [typeRaw, ...rest] = raw.trim().split(/\s+/);
  const type = typeRaw.toUpperCase().replace("_ARROW", "");
  const isArrow = typeRaw.endsWith("_ARROW");
  const data = rest.join(" ");
  const svgHeight = height;
  const svgWidth = width;
  const sw = strokeWidthPx && strokeWidthPx > 0 ? strokeWidthPx : 2;
  const rotationStyle = rotationDeg ? `transform: rotate(${rotationDeg}deg); transform-origin: center;` : "";
  const dashAttr = dashArray ? `stroke-dasharray="${dashArray}"` : "";
  const commonStyle = `
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${svgWidth}px;
    height: ${svgHeight}px;
    ${rotationStyle}
  `;
  switch (type) {
    case "PATH": {
      const defs = buildMarkerDefs(headEnd, tailEnd, strokeColorOpt || "#000");
      const markerStartAttr = defs.startId ? `marker-start="url(#${defs.startId})"` : "";
      const markerEndAttr = defs.endId ? `marker-end="url(#${defs.endId})"` : "";
      return `<svg viewBox="0 0 100 100" style="${commonStyle}" overflow="visible">
        ${defs.defs}
        <path d="${data}" fill="none" stroke="${strokeColorOpt || "#000"}" stroke-width="${sw}" ${dashAttr} ${scaleStrokes ? "" : 'vector-effect="non-scaling-stroke"'} ${markerStartAttr} ${markerEndAttr} />
      </svg>`;
    }
    case "POLYLINE":
    case "LINE": {
      const coords = data.split(/[\s,]+/).map((v) => parseFloat(v)).filter((v) => !isNaN(v));
      if (coords.length < 4 || coords.length % 2 !== 0) {
        if (typeof console !== "undefined" && console.warn) {
          console.warn(`[pptx-to-html] Invalid POLYLINE/LINE shape data: "${data}"`);
        }
        return "";
      }
      const pointPairs = [];
      for (let i = 0; i < coords.length; i += 2) {
        pointPairs.push(`${coords[i]},${coords[i + 1]}`);
      }
      const effectiveWidth = width > 0 ? width : Math.max(sw * 2, 2);
      const effectiveHeight = height > 0 ? height : Math.max(sw * 2, 2);
      const defs = buildMarkerDefs(headEnd, tailEnd ?? (isArrow ? { type: "triangle", w: "med", len: "med" } : void 0), strokeColorOpt || "#000");
      const markerStartAttr = defs.startId ? `marker-start="url(#${defs.startId})"` : "";
      const markerEndAttr = defs.endId ? `marker-end="url(#${defs.endId})"` : "";
      const scaledPairs = [];
      for (let i = 0; i < coords.length; i += 2) {
        const px = coords[i] / 100 * effectiveWidth;
        const py = coords[i + 1] / 100 * effectiveHeight;
        scaledPairs.push(`${px},${py}`);
      }
      const scaledPoints = scaledPairs.join(" ");
      return `
        <svg viewBox="0 0 ${effectiveWidth} ${effectiveHeight}"
            style="
              position: absolute;
              left: ${x}px;
              top: ${y}px;
              width: ${effectiveWidth}px;
              height: ${effectiveHeight}px;
              ${rotationStyle}
            "
            overflow="visible">
          ${defs.defs}
          <polyline points="${scaledPoints}"
                    fill="none"
                    stroke="${strokeColorOpt || "#000"}"
                    stroke-width="${sw}"
                    ${dashAttr}
                    ${scaleStrokes ? "" : 'vector-effect="non-scaling-stroke"'}
                    ${markerStartAttr} ${markerEndAttr} />
        </svg>`;
    }
    case "POLYGON":
    default:
      return `<svg viewBox="0 0 100 100" style="${commonStyle}">
        <polygon points="${data}" fill="${fill}" stroke="${strokeColorOpt ?? "none"}" stroke-width="${sw}" ${dashAttr} ${scaleStrokes ? "" : 'vector-effect="non-scaling-stroke"'} />
      </svg>`;
  }
}
function buildMarkerDefs(headEnd, tailEnd, color) {
  const parts = [];
  let startId;
  let endId;
  let startLen = 0;
  let endLen = 0;
  if (headEnd && headEnd.type && headEnd.type !== "none") {
    startId = `mstart-${Math.random().toString(36).slice(2, 8)}`;
    const info = markerDef(startId, headEnd, color);
    parts.push(info.svg);
    startLen = info.len;
  }
  if (tailEnd && tailEnd.type && tailEnd.type !== "none") {
    endId = `mend-${Math.random().toString(36).slice(2, 8)}`;
    const info = markerDef(endId, tailEnd, color);
    parts.push(info.svg);
    endLen = info.len;
  }
  return { defs: parts.length ? `<defs>${parts.join("\n")}</defs>` : "", startId, endId, startLen, endLen };
}
function markerDef(id, spec, color) {
  const lenFactor = mapSize(spec.len);
  const wFactor = mapSize(spec.w);
  const len = 3 * lenFactor;
  const w = 2.5 * wFactor;
  const refX = 0;
  const refY = w / 2;
  let svg;
  switch ((spec.type || "triangle").toLowerCase()) {
    case "diamond":
      svg = `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${len}" markerHeight="${w}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <polygon points="${len / 2},0 ${len},${w / 2} ${len / 2},${w} 0,${w / 2}" fill="${color}" />
              </marker>`;
      break;
    case "oval":
      svg = `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${len}" markerHeight="${w}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <ellipse cx="${len / 2}" cy="${w / 2}" rx="${len / 2}" ry="${w / 2}" fill="${color}" />
              </marker>`;
      break;
    case "stealth": {
      const notch = len * 0.35;
      svg = `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${len}" markerHeight="${w}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <polygon points="0,0 ${len},${w / 2} 0,${w} ${notch},${w / 2}" fill="${color}" />
              </marker>`;
      break;
    }
    case "arrow":
    case "triangle":
    default:
      svg = `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${len}" markerHeight="${w}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <polygon points="0,0 ${len},${w / 2} 0,${w}" fill="${color}" />
              </marker>`;
      break;
  }
  return { svg, len };
}
function mapSize(size) {
  switch ((size || "med").toLowerCase()) {
    case "sm":
    case "small":
      return 0.7;
    case "lg":
    case "large":
      return 1.4;
    case "med":
    case "medium":
    default:
      return 1;
  }
}

// src/renderer/renderTableElement.ts
function renderTableElement(el) {
  const nf = (n, fb = 0) => Number.isFinite(n) ? n : fb;
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const width = nf(el.size?.width, 0) / 9525;
  const height = nf(el.size?.height, 0) / 9525;
  const colWidthsPx = el.columns.map((w) => nf(w, 0) / 9525);
  const colTotal = colWidthsPx.reduce((a, b) => a + b, 0) || 1;
  const cols = colWidthsPx.map((w) => `<col style="width:${w / colTotal * 100}%">`).join("");
  const tableBg = el.tableFillColor || el.style?.fills?.wholeTbl;
  let rowIndex = 0;
  const rowsHtml = el.rows.map((row) => {
    let colIndex = 0;
    const tds = row.cells.map((cell) => {
      const pad = cell.padding || { left: 6, top: 2, right: 6, bottom: 2 };
      const ta = cell.align?.horizontal || "left";
      const va = cell.align?.vertical || "top";
      const borderCss = computeCellBordersCSS(el, cell, rowIndex, colIndex);
      const { bg, fontColor, emphasize } = computeCellStyleFromTableStyle(el, rowIndex, colIndex);
      const isHeaderCol = emphasize;
      const style = `
            padding:${pad.top}px ${pad.right}px ${pad.bottom}px ${pad.left}px;
            text-align:${ta};
            vertical-align:${va === "middle" ? "middle" : va};
            ${cell.fillColor ? `background-color:${cell.fillColor};` : ""}
            ${!cell.fillColor && bg ? `background-color:${bg};` : ""}
            ${cell.font?.color ? `color:${cell.font.color};` : fontColor ? `color:${fontColor};` : ""}
            ${cell.font?.name ? `font-family:${cell.font.name};` : ""}
            ${cell.font?.size ? `font-size:${cell.font.size}pt;` : ""}
            ${borderCss}
            ${isHeaderCol ? "font-weight:600;" : ""}
            overflow:hidden; word-break: break-word; white-space: pre-wrap;`;
      const span = `${cell.colSpan ? ` colspan="${cell.colSpan}"` : ""}${cell.rowSpan ? ` rowspan="${cell.rowSpan}"` : ""}`;
      const content = escape2(cell.text).replace(/\n/g, "<br>");
      const html = `<td${span} style="${style}">${content}</td>`;
      colIndex += cell.colSpan || 1;
      return html;
    }).join("");
    const rowStyle = el.tableStyle?.firstRow && rowIndex === 0 ? ' style="font-weight:600;"' : "";
    const trHtml = `<tr${rowStyle}>${tds}</tr>`;
    rowIndex += 1;
    return trHtml;
  }).join("");
  return `<div style="position:absolute; left:${x}px; top:${y}px; width:${width}px; height:${height}px;">
    <table style="border-collapse:collapse; width:100%; height:100%; table-layout:fixed;${tableBg ? ` background-color:${tableBg};` : ""}">
      <colgroup>${cols}</colgroup>
      <tbody>${rowsHtml}</tbody>
    </table>
  </div>`;
}
function computeCellBordersCSS(el, cell, rowIndex, colIndex) {
  const css = [];
  const sides = ["top", "right", "bottom", "left"];
  const apply = (side, b) => {
    if (!b) return;
    const w = b.width ?? 1;
    const c = b.color ?? "#000";
    const st = b.style === "dashed" || b.style === "dotted" ? b.style : "solid";
    css.push(`border-${side}: ${Math.max(1, Math.round(w))}px ${st} ${c};`);
  };
  for (const s of sides) {
    const b = cell.borders?.[s];
    if (b) apply(s, b);
  }
  const tb = el.tableBorders || {};
  const lastRow = rowIndex === el.rows.length - 1;
  const lastCol = colIndex === el.columns.length - 1;
  if (!cell.borders?.top && rowIndex === 0) apply("top", tb.top);
  if (!cell.borders?.bottom && lastRow) apply("bottom", tb.bottom);
  if (!cell.borders?.left && colIndex === 0) apply("left", tb.left);
  if (!cell.borders?.right && lastCol) apply("right", tb.right);
  if (!cell.borders?.top && rowIndex > 0) apply("top", tb.insideH);
  if (!cell.borders?.left && colIndex > 0) apply("left", tb.insideV);
  return css.join(" ");
}
function escape2(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function computeCellStyleFromTableStyle(el, rowIndex, colIndex) {
  const lastRow = rowIndex === el.rows.length - 1;
  const lastCol = colIndex === el.columns.length - 1;
  const s = el.style || {};
  const fills = s.fills || {};
  const fontColors = s.fontColors || {};
  let emphasize = false;
  if (el.tableStyle?.firstRow && rowIndex === 0) {
    emphasize = true;
    return { bg: fills.firstRow || fills.wholeTbl || "#f0f0f0", fontColor: fontColors.firstRow || fontColors.wholeTbl, emphasize };
  }
  if (el.tableStyle?.lastRow && lastRow) {
    return { bg: fills.lastRow || fills.wholeTbl || "#f0f0f0", fontColor: fontColors.lastRow || fontColors.wholeTbl, emphasize };
  }
  if (el.tableStyle?.firstCol && colIndex === 0) {
    emphasize = true;
    return { bg: fills.firstCol || fills.wholeTbl || "#f0f0f0", fontColor: fontColors.firstCol || fontColors.wholeTbl, emphasize };
  }
  if (el.tableStyle?.lastCol && lastCol) {
    return { bg: fills.lastCol || fills.wholeTbl || "#f0f0f0", fontColor: fontColors.lastCol || fontColors.wholeTbl, emphasize };
  }
  if (el.tableStyle?.bandRow) {
    const baseIndex = el.tableStyle?.firstRow ? rowIndex - 1 : rowIndex;
    const band = baseIndex % 2 === 0 ? "band1H" : "band2H";
    return { bg: fills[band] || fills.wholeTbl || (rowIndex % 2 === 1 ? "#fafafa" : void 0), fontColor: fontColors[band] || fontColors.wholeTbl, emphasize };
  }
  if (el.tableStyle?.bandCol) {
    const band = colIndex % 2 === 0 ? "band1V" : "band2V";
    return { bg: fills[band] || fills.wholeTbl || (colIndex % 2 === 1 ? "#fafafa" : void 0), fontColor: fontColors[band] || fontColors.wholeTbl, emphasize };
  }
  return { bg: fills.wholeTbl, fontColor: fontColors.wholeTbl, emphasize };
}

// src/renderer/renderChartElement.ts
function renderChartElement(el) {
  const nf = (n, fb = 0) => Number.isFinite(n) ? n : fb;
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const width = Math.max(1, nf(el.size?.width, 0) / 9525);
  const height = Math.max(1, nf(el.size?.height, 0) / 9525);
  const padding = 24;
  const palette = el.palette && el.palette.length > 0 ? el.palette : ["#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc949"];
  let svg = "";
  if (el.chartType === "column" || el.chartType === "bar") {
    svg = renderBarLike(el, width, height, padding, palette);
  } else if (el.chartType === "line") {
    svg = renderLine(el, width, height, padding, palette);
  } else if (el.chartType === "area") {
    svg = renderArea(el, width, height, padding, palette);
  } else if (el.chartType === "pie") {
    svg = renderPie(el, width, height, palette);
  } else if (el.chartType === "scatter") {
    svg = renderScatter(el, width, height, padding, palette);
  }
  const title = el.title ? `<div style="position:absolute;left:${x}px;top:${y - 20}px;width:${width}px;text-align:center;font-weight:600;">${escape3(el.title)}</div>` : "";
  return `${title}<div style="position:absolute; left:${x}px; top:${y}px; width:${width}px; height:${height}px;">
      <svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
        ${svg}
      </svg>
    </div>`;
}
function renderBarLike(el, width, height, pad, palette) {
  const catCount = el.categories.length || 1;
  const chartW = width - pad * 2;
  const chartH = height - pad * 2;
  const isHorizontal = el.chartType === "bar";
  const stacked = el.stackedMode && el.stackedMode !== "none";
  const percent = el.stackedMode === "percent";
  let maxVal = 1;
  if (stacked) {
    const sums = new Array(catCount).fill(0).map((_, i) => el.series.reduce((acc, s) => acc + ((s.values || [])[i] || 0), 0));
    maxVal = percent ? 1 : Math.max(1, ...sums);
  } else {
    let mv = 1;
    for (const s of el.series) {
      const vals = s.values || [];
      for (const v of vals) mv = Math.max(mv, v);
    }
    maxVal = mv;
  }
  const parts = [];
  if (isHorizontal) {
    parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
    parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
    const ticks = computeTicks(0, maxVal, 4);
    ticks.forEach((t) => {
      const tx = pad + t / maxVal * chartW;
      const ty = height - pad;
      parts.push(`<line x1="${tx}" y1="${ty}" x2="${tx}" y2="${ty + 4}" stroke="#999" stroke-width="1" />`);
      parts.push(`<text x="${tx}" y="${ty + 16}" text-anchor="middle" font-size="10" fill="#666">${formatNumber(t, percent ? "0%" : el.valueFormat)}</text>`);
    });
  } else {
    parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
    parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
    const ticks = computeTicks(0, maxVal, 4);
    ticks.forEach((t) => {
      const ty = height - pad - t / maxVal * chartH;
      parts.push(`<line x1="${pad - 4}" y1="${ty}" x2="${pad}" y2="${ty}" stroke="#999" stroke-width="1" />`);
      parts.push(`<text x="${pad - 6}" y="${ty + 3}" text-anchor="end" font-size="10" fill="#666">${formatNumber(t, percent ? "0%" : el.valueFormat)}</text>`);
      parts.push(`<line x1="${pad}" y1="${ty}" x2="${width - pad}" y2="${ty}" stroke="#eee" stroke-width="1" />`);
    });
  }
  if (isHorizontal) {
    const catBand = chartH / catCount;
    const barH = Math.max(2, catBand * 0.8 / (stacked ? 1 : el.series.length));
    el.series.forEach((s, si) => {
      const color = s.color || palette[si % palette.length];
      (s.values || []).forEach((v, ci) => {
        const baseY = pad + ci * catBand + (catBand - (stacked ? barH : el.series.length * barH)) / 2;
        if (stacked) {
          const prev = el.series.slice(0, si).reduce((acc, ss) => acc + ((ss.values || [])[ci] || 0), 0);
          const sum = el.series.reduce((acc, ss) => acc + ((ss.values || [])[ci] || 0), 0) || 1;
          const start = (percent ? prev / sum : prev) / maxVal * chartW;
          const w = (percent ? (v || 0) / sum : v) / maxVal * chartW;
          const x = pad + start;
          const y = baseY;
          parts.push(`<rect x="${x}" y="${y}" width="${w}" height="${barH}" fill="${color}" />`);
          if (el.showDataLabels) {
            const fmt = s.valueFormat || el.valueFormat;
            parts.push(`<text x="${x + w - 2}" y="${y + barH / 2 + 3}" text-anchor="end" font-size="10" fill="#000">${formatNumber(percent ? v / sum : v, percent ? "0%" : fmt)}</text>`);
          }
        } else {
          const w = v / maxVal * chartW;
          const y = baseY + si * barH;
          parts.push(`<rect x="${pad}" y="${y}" width="${w}" height="${barH}" fill="${color}" />`);
          if (el.showDataLabels) {
            const fmt = s.valueFormat || el.valueFormat;
            parts.push(`<text x="${pad + w + 2}" y="${y + barH / 2 + 3}" font-size="10" fill="#000">${formatNumber(v, fmt)}</text>`);
          }
        }
      });
    });
    const labelSize = Math.max(10, Math.min(12, chartH / (catCount * 2)));
    el.categories.forEach((c, i) => {
      const cy = pad + i * catBand + catBand / 2 + 4;
      parts.push(`<text x="${pad - 8}" y="${cy}" text-anchor="end" font-size="${labelSize}" fill="#333">${escape3(String(c))}</text>`);
    });
  } else {
    const catBand = chartW / catCount;
    const barW = Math.max(2, catBand * 0.8 / (stacked ? 1 : el.series.length));
    el.series.forEach((s, si) => {
      const color = s.color || palette[si % palette.length];
      (s.values || []).forEach((v, ci) => {
        const baseX = pad + ci * catBand + (catBand - (stacked ? barW : el.series.length * barW)) / 2;
        if (stacked) {
          const prev = el.series.slice(0, si).reduce((acc, ss) => acc + ((ss.values || [])[ci] || 0), 0);
          const sum = el.series.reduce((acc, ss) => acc + ((ss.values || [])[ci] || 0), 0) || 1;
          const start = (percent ? prev / sum : prev) / maxVal * chartH;
          const h = (percent ? (v || 0) / sum : v) / maxVal * chartH;
          const x = baseX;
          const y = height - pad - start - h;
          parts.push(`<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="${color}" />`);
          if (el.showDataLabels) {
            const fmt = s.valueFormat || el.valueFormat;
            parts.push(`<text x="${x + barW / 2}" y="${y - 2}" text-anchor="middle" font-size="10" fill="#000">${formatNumber(percent ? v / sum : v, percent ? "0%" : fmt)}</text>`);
          }
        } else {
          const h = v / maxVal * chartH;
          const x = baseX + si * barW;
          const y = height - pad - h;
          parts.push(`<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="${color}" />`);
          if (el.showDataLabels) {
            const fmt = s.valueFormat || el.valueFormat;
            parts.push(`<text x="${x + barW / 2}" y="${y - 2}" text-anchor="middle" font-size="10" fill="#000">${formatNumber(v, fmt)}</text>`);
          }
        }
      });
    });
    const labelSize = Math.max(10, Math.min(12, chartW / (catCount * 4)));
    el.categories.forEach((c, i) => {
      const cx = pad + i * catBand + catBand / 2;
      const cy = height - pad + 14;
      parts.push(`<text x="${cx}" y="${cy}" text-anchor="middle" font-size="${labelSize}" fill="#333">${escape3(String(c))}</text>`);
    });
  }
  if (el.showLegend) {
    parts.push(renderLegend(el, width, pad, palette));
  }
  return parts.join("\n");
}
function renderLine(el, width, height, pad, palette) {
  const catCount = el.categories.length || 1;
  const chartW = width - pad * 2;
  const chartH = height - pad * 2;
  let maxVal = 1;
  for (const s of el.series) {
    const vals = s.values || [];
    for (const v of vals) maxVal = Math.max(maxVal, v);
  }
  const xStep = chartW / Math.max(1, catCount - 1);
  const parts = [];
  parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  const ticks = computeTicks(0, maxVal, 4);
  ticks.forEach((t) => {
    const ty = height - pad - t / maxVal * chartH;
    parts.push(`<line x1="${pad - 4}" y1="${ty}" x2="${pad}" y2="${ty}" stroke="#999" stroke-width="1" />`);
    parts.push(`<text x="${pad - 6}" y="${ty + 3}" text-anchor="end" font-size="10" fill="#666">${formatNumber(t, el.valueFormat)}</text>`);
    parts.push(`<line x1="${pad}" y1="${ty}" x2="${width - pad}" y2="${ty}" stroke="#eee" stroke-width="1" />`);
  });
  const stacked = el.stackedMode && el.stackedMode !== "none";
  const percent = el.stackedMode === "percent";
  const totals = percent ? new Array(catCount).fill(0).map((_, i) => el.series.reduce((acc, sr) => acc + ((sr.values || [])[i] || 0), 0)) : void 0;
  el.series.forEach((s, si) => {
    const color = s.color || palette[si % palette.length];
    let d = "";
    (s.values || []).forEach((v, i) => {
      const x = pad + i * xStep;
      let val = v;
      if (stacked) {
        const prev = el.series.slice(0, si).reduce((acc, ss) => acc + ((ss.values || [])[i] || 0), 0);
        val = prev + v;
        if (percent && totals) val = totals[i] ? val / totals[i] : 0;
      }
      const y = height - pad - val / (percent ? 1 : maxVal) * chartH;
      d += i === 0 ? `M ${x} ${y}` : ` L ${x} ${y}`;
    });
    parts.push(`<path d="${d}" fill="none" stroke="${color}" stroke-width="2" />`);
    if (el.showDataLabels) {
      (s.values || []).forEach((v, i) => {
        const x = pad + i * xStep;
        const basePrev = stacked ? el.series.slice(0, si).reduce((acc, ss) => acc + ((ss.values || [])[i] || 0), 0) : 0;
        const dispVal = percent && totals ? v / (totals[i] || 1) : v;
        const y = height - pad - (stacked ? basePrev + v : v) / (percent ? 1 : maxVal) * chartH;
        parts.push(`<circle cx="${x}" cy="${y}" r="2.5" fill="${color}" />`);
        const fmt = s.valueFormat || el.valueFormat;
        parts.push(`<text x="${x}" y="${y - 6}" text-anchor="middle" font-size="10" fill="#000">${formatNumber(dispVal, percent ? "0%" : fmt)}</text>`);
      });
    }
  });
  return parts.join("\n");
}
function renderArea(el, width, height, pad, palette) {
  const catCount = el.categories.length || 1;
  const chartW = width - pad * 2;
  const chartH = height - pad * 2;
  let maxVal = 1;
  for (const s of el.series) {
    const vals = s.values || [];
    for (const v of vals) maxVal = Math.max(maxVal, v);
  }
  const xStep = chartW / Math.max(1, catCount - 1);
  const parts = [];
  parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  const stacked = el.stackedMode && el.stackedMode !== "none";
  const percent = el.stackedMode === "percent";
  const totals = percent ? new Array(catCount).fill(0).map((_, i) => el.series.reduce((acc, sr) => acc + ((sr.values || [])[i] || 0), 0)) : void 0;
  const baseline = new Array(catCount).fill(0);
  el.series.forEach((s, si) => {
    const color = s.color || palette[si % palette.length];
    const topY = [];
    const botY = [];
    (s.values || []).forEach((v, i) => {
      const prev = stacked ? baseline[i] : 0;
      const val = stacked ? prev + (percent && totals ? totals[i] ? v / totals[i] : 0 : v) : v;
      const top = height - pad - val / (percent ? 1 : maxVal) * chartH;
      const bottom = height - pad - prev / (percent ? 1 : maxVal) * chartH;
      topY.push(top);
      botY.push(bottom);
      if (stacked) baseline[i] = percent && totals ? val : prev + v;
    });
    let d = "";
    for (let i = 0; i < topY.length; i++) {
      const x = pad + i * xStep;
      d += i === 0 ? `M ${x} ${topY[i]}` : ` L ${x} ${topY[i]}`;
    }
    for (let i = botY.length - 1; i >= 0; i--) {
      const x = pad + i * xStep;
      d += ` L ${x} ${botY[i]}`;
    }
    d += " Z";
    parts.push(`<path d="${d}" fill="${color}" fill-opacity="0.6" stroke="none" />`);
  });
  return parts.join("\n");
}
function renderPie(el, width, height, palette) {
  const s0 = el.series[0];
  const values = s0 && s0.values ? s0.values : [];
  const total = values.reduce((a, b) => a + Math.max(0, b), 0) || 1;
  const cx = width / 2;
  const cy = height / 2;
  const r = Math.min(width, height) * 0.35;
  let start = 0;
  const parts = [];
  values.forEach((v, i) => {
    const frac = Math.max(0, v) / total;
    const end = start + frac * 2 * Math.PI;
    const x1 = cx + r * Math.cos(start);
    const y1 = cy + r * Math.sin(start);
    const x2 = cx + r * Math.cos(end);
    const y2 = cy + r * Math.sin(end);
    const large = end - start > Math.PI ? 1 : 0;
    const color = s0 && s0.color || palette[i % palette.length];
    const d = `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`;
    parts.push(`<path d="${d}" fill="${color}" />`);
    if (el.showDataLabels && frac > 0) {
      const mid = (start + end) / 2;
      const lx = cx + (r + 12) * Math.cos(mid);
      const ly = cy + (r + 12) * Math.sin(mid);
      parts.push(`<text x="${lx}" y="${ly}" text-anchor="middle" font-size="10" fill="#000">${(frac * 100).toFixed(0)}%</text>`);
    }
    start = end;
  });
  return parts.join("\n");
}
function renderScatter(el, width, height, pad, palette) {
  const chartW = width - pad * 2;
  const chartH = height - pad * 2;
  const allPoints = [];
  for (const s of el.series) {
    const pts = s.points || [];
    for (const p of pts) allPoints.push(p);
  }
  const minX = Math.min(...allPoints.map((p) => p.x), 0);
  const maxX = Math.max(...allPoints.map((p) => p.x), 1);
  const minY = Math.min(...allPoints.map((p) => p.y), 0);
  const maxY = Math.max(...allPoints.map((p) => p.y), 1);
  const parts = [];
  parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  const xticks = computeTicks(minX, maxX, 4);
  xticks.forEach((t) => {
    const tx = pad + (t - minX) / Math.max(1e-9, maxX - minX) * chartW;
    const ty = height - pad;
    parts.push(`<line x1="${tx}" y1="${ty}" x2="${tx}" y2="${ty + 4}" stroke="#999" stroke-width="1" />`);
    parts.push(`<text x="${tx}" y="${ty + 16}" text-anchor="middle" font-size="10" fill="#666">${formatNumber(t, el.valueFormat)}</text>`);
  });
  const yticks = computeTicks(minY, maxY, 4);
  yticks.forEach((t) => {
    const ty = height - pad - (t - minY) / Math.max(1e-9, maxY - minY) * chartH;
    parts.push(`<line x1="${pad - 4}" y1="${ty}" x2="${pad}" y2="${ty}" stroke="#999" stroke-width="1" />`);
    parts.push(`<text x="${pad - 6}" y="${ty + 3}" text-anchor="end" font-size="10" fill="#666">${formatNumber(t, el.valueFormat)}</text>`);
    parts.push(`<line x1="${pad}" y1="${ty}" x2="${width - pad}" y2="${ty}" stroke="#eee" stroke-width="1" />`);
  });
  el.series.forEach((s, si) => {
    const color = s.color || palette[si % palette.length];
    (s.points || []).forEach((p) => {
      const x = pad + (p.x - minX) / Math.max(1e-9, maxX - minX) * chartW;
      const y = height - pad - (p.y - minY) / Math.max(1e-9, maxY - minY) * chartH;
      parts.push(`<circle cx="${x}" cy="${y}" r="3" fill="${color}" />`);
      if (el.showDataLabels) {
        const fmt = s.valueFormat || el.valueFormat;
        parts.push(`<text x="${x + 5}" y="${y - 5}" font-size="10" fill="#000">${formatNumber(p.y, fmt)}</text>`);
      }
    });
  });
  return parts.join("\n");
}
function escape3(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function computeTicks(min, max, count) {
  const span = max - min;
  if (span <= 0) return [min, max];
  const step = niceNum(span / count, true);
  const ticks = [];
  let v = Math.ceil(min / step) * step;
  while (v <= max + 1e-9) {
    ticks.push(Number(v.toFixed(10)));
    v += step;
  }
  return ticks;
}
function niceNum(range, round) {
  const exponent = Math.floor(Math.log10(range));
  const fraction = range / Math.pow(10, exponent);
  let niceFraction;
  if (round) {
    if (fraction < 1.5) niceFraction = 1;
    else if (fraction < 3) niceFraction = 2;
    else if (fraction < 7) niceFraction = 5;
    else niceFraction = 10;
  } else {
    if (fraction <= 1) niceFraction = 1;
    else if (fraction <= 2) niceFraction = 2;
    else if (fraction <= 5) niceFraction = 5;
    else niceFraction = 10;
  }
  return niceFraction * Math.pow(10, exponent);
}
function formatNumber(v, formatCode) {
  if (!formatCode) {
    if (Math.abs(v) >= 1e3) return `${Math.round(v)}`;
    if (Math.abs(v) >= 10) return v.toFixed(0);
    if (Math.abs(v) >= 1) return v.toFixed(1);
    return v.toFixed(2);
  }
  let isPercent = /%/.test(formatCode);
  let decimals = 0;
  const decMatch = formatCode.match(/\.([0#]+)/);
  if (decMatch) decimals = decMatch[1].length;
  const currencyMatch = formatCode.match(/([$€£¥])/);
  const currency = currencyMatch ? currencyMatch[1] : "";
  const useThousands = /#,##0/.test(formatCode);
  let n = v;
  if (isPercent) n = n * 100;
  let str = n.toFixed(decimals);
  if (useThousands) {
    const [int, frac] = str.split(".");
    const withSep = int.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    str = frac !== void 0 ? `${withSep}.${frac}` : withSep;
  }
  if (currency) str = currency + str;
  if (isPercent) str = str + "%";
  return str;
}
function renderLegend(el, width, pad, palette) {
  let x = pad;
  const y = pad - 12;
  const parts = [];
  el.series.forEach((s, i) => {
    const color = s.color || palette[i % palette.length];
    const label = s.name || `Series ${i + 1}`;
    parts.push(`<rect x="${x}" y="${y}" width="10" height="10" fill="${color}" />`);
    parts.push(`<text x="${x + 14}" y="${y + 9}" font-size="10" fill="#333">${escape3(label)}</text>`);
    x += 14 + label.length * 6 + 10;
  });
  return parts.join("\n");
}

// src/renderer/HtmlRenderer.ts
function backgroundFillToCss(fill) {
  switch (fill.type) {
    case "solid":
      return `background-color: ${fill.color};`;
    case "gradient": {
      if (fill.stops.length === 0) return "background-color: transparent;";
      if (fill.stops.length === 1) return `background-color: ${fill.stops[0].color};`;
      const stops = fill.stops.map((s) => `${s.color} ${s.offset}%`).join(", ");
      if (fill.gradientType === "radial") {
        return `background: radial-gradient(ellipse at center, ${stops});`;
      }
      const angle = fill.angle ?? 180;
      return `background: linear-gradient(${angle}deg, ${stops});`;
    }
    case "image":
      return `background-image: url('${fill.src}'); background-size: cover; background-position: center; background-repeat: no-repeat;`;
    case "none":
      return "background-color: transparent;";
  }
}
var HtmlRenderer = class {
  /**
   * Renders a slide to an HTML <div> with all elements positioned accordingly.
   * @param elements List of SlideElement (text, image, shape)
   * @param options Optional width and height (in px) for the slide container.
   *  - If scaleToFit=true, width/height define the outer container size and contents are scaled from the base 960x540.
   * @returns HTML string representing the slide
   */
  static render(elements, options = {}) {
    const baseW = options.baseWidth ?? 960;
    const baseH = options.baseHeight ?? 540;
    const targetW = options.width ?? baseW;
    const targetH = options.height ?? baseH;
    const scaleToFit = options.scaleToFit === true;
    const letterbox = scaleToFit ? options.letterbox !== false : options.letterbox === true;
    const htmlParts = elements.map((el) => {
      switch (el.type) {
        case "background": {
          const bgEl = el;
          const styleBg = backgroundFillToCss(bgEl.fill);
          return `<div style="position:absolute; left:0; top:0; width:${baseW}px; height:${baseH}px; ${styleBg}"></div>`;
        }
        case "text":
          return renderTextElement(el);
        case "image":
          return renderImageElement(el);
        case "shape":
          return renderShapeElement(el, { scaleStrokes: scaleToFit });
        case "table":
          return renderTableElement(el);
        case "chart":
          return renderChartElement(el);
        default:
          if (typeof console !== "undefined" && console.warn) {
            console.warn(`[pptx-to-html] Unsupported element type: ${el?.type}`);
          }
          return "";
      }
    });
    if (scaleToFit) {
      if (letterbox) {
        const s = Math.min(targetW / baseW, targetH / baseH);
        const offsetX = (targetW - baseW * s) / 2;
        const offsetY = (targetH - baseH * s) / 2;
        return `<div class="slide-container" style="position: relative; width: ${targetW}px; height: ${targetH}px; overflow: hidden; background-color: #000;">
            <div class="slide" style="position: absolute; left: ${offsetX}px; top: ${offsetY}px; width: ${baseW}px; height: ${baseH}px; transform: scale(${s}); transform-origin: top left; background-color: #fff;">
              ${htmlParts.join("\n")}
            </div>
          </div>`;
      } else {
        const sx = targetW / baseW;
        const sy = targetH / baseH;
        return `<div class="slide-container" style="position: relative; width: ${targetW}px; height: ${targetH}px; overflow: hidden;">
            <div class="slide" style="position: absolute; left: 0; top: 0; width: ${baseW}px; height: ${baseH}px; transform: scale(${sx}, ${sy}); transform-origin: top left; background-color: #fff;">
              ${htmlParts.join("\n")}
            </div>
          </div>`;
      }
    }
    return `<div class="slide" style="position: relative; width: ${targetW}px; height: ${targetH}px; overflow: hidden; background-color: #fff;">
        ${htmlParts.join("\n")}
      </div>`;
  }
};

// src/index.ts
async function pptxToHtml(buffer, config) {
  if (config?.domParserFactory) {
    const { XmlHelper: XmlHelper2 } = await import("./XmlHelper-47TKKPUJ.js");
    XmlHelper2.setDomParser(config.domParserFactory);
  }
  const reader = new PptxReader();
  const slides = await reader.load(buffer);
  const base = await reader.getBaseSizePx();
  const opts = { ...config || {}, baseWidth: base.width, baseHeight: base.height };
  return slides.map((slideElements) => HtmlRenderer.render(slideElements, opts));
}
export {
  pptxToHtml
};
//# sourceMappingURL=index.js.map