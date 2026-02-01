var __require = /* @__PURE__ */ ((x) => typeof require !== "undefined" ? require : typeof Proxy !== "undefined" ? new Proxy(x, {
  get: (a, b) => (typeof require !== "undefined" ? require : a)[b]
}) : x)(function(x) {
  if (typeof require !== "undefined") return require.apply(this, arguments);
  throw Error('Dynamic require of "' + x + '" is not supported');
});

// src/core/XmlHelper.ts
function libWarn(msg) {
  if (typeof console !== "undefined" && console.warn) {
    console.warn(`[pptx-to-html] ${msg}`);
  }
}
function hexToRgb(hex) {
  const h = hex.replace("#", "");
  return [
    parseInt(h.substring(0, 2), 16),
    parseInt(h.substring(2, 4), 16),
    parseInt(h.substring(4, 6), 16)
  ];
}
function rgbToHex(r, g, b) {
  const clamp = (v) => Math.max(0, Math.min(255, Math.round(v)));
  return `#${[clamp(r), clamp(g), clamp(b)].map((v) => v.toString(16).padStart(2, "0")).join("")}`;
}
function rgbToHsl(r, g, b) {
  const rn = r / 255, gn = g / 255, bn = b / 255;
  const max = Math.max(rn, gn, bn), min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = 0;
  if (max === rn) h = ((gn - bn) / d + (gn < bn ? 6 : 0)) / 6;
  else if (max === gn) h = ((bn - rn) / d + 2) / 6;
  else h = ((rn - gn) / d + 4) / 6;
  return [h * 360, s, l];
}
function hslToRgb(h, s, l) {
  if (s === 0) {
    const v = Math.round(l * 255);
    return [v, v, v];
  }
  const hue2rgb = (p2, q2, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p2 + (q2 - p2) * 6 * t;
    if (t < 1 / 2) return q2;
    if (t < 2 / 3) return p2 + (q2 - p2) * (2 / 3 - t) * 6;
    return p2;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hn = h / 360;
  return [
    Math.round(hue2rgb(p, q, hn + 1 / 3) * 255),
    Math.round(hue2rgb(p, q, hn) * 255),
    Math.round(hue2rgb(p, q, hn - 1 / 3) * 255)
  ];
}
function applyColorModifiers(hex, colorEl) {
  const children = Array.from(colorEl.children);
  if (children.length === 0) return hex;
  let [r, g, b] = hexToRgb(hex);
  let tint;
  let shade;
  let lumMod;
  let lumOff;
  let satMod;
  let satOff;
  for (const child of children) {
    const tag = child.localName;
    const val = Number(child.getAttribute("val") || "0");
    if (!Number.isFinite(val)) continue;
    switch (tag) {
      case "tint":
        tint = val;
        break;
      case "shade":
        shade = val;
        break;
      case "lumMod":
        lumMod = val;
        break;
      case "lumOff":
        lumOff = val;
        break;
      case "satMod":
        satMod = val;
        break;
      case "satOff":
        satOff = val;
        break;
    }
  }
  if (tint !== void 0) {
    const t = tint / 1e5;
    r = Math.round(r + (255 - r) * t);
    g = Math.round(g + (255 - g) * t);
    b = Math.round(b + (255 - b) * t);
  }
  if (shade !== void 0) {
    const s = shade / 1e5;
    r = Math.round(r * s);
    g = Math.round(g * s);
    b = Math.round(b * s);
  }
  if (lumMod !== void 0 || lumOff !== void 0 || satMod !== void 0 || satOff !== void 0) {
    let [h, s, l] = rgbToHsl(r, g, b);
    if (lumMod !== void 0) l = l * (lumMod / 1e5);
    if (lumOff !== void 0) l = l + lumOff / 1e5;
    if (satMod !== void 0) s = s * (satMod / 1e5);
    if (satOff !== void 0) s = s + satOff / 1e5;
    l = Math.max(0, Math.min(1, l));
    s = Math.max(0, Math.min(1, s));
    [r, g, b] = hslToRgb(h, s, l);
  }
  return rgbToHex(r, g, b);
}
var XmlHelper = class _XmlHelper {
  static domParserFactory = null;
  /**
   * Parses a string containing XML into a DOM Document
   * @param xmlString XML string to parse
   * @returns DOM Document
   */
  static parseXml(xmlString) {
    if (_XmlHelper.domParserFactory) {
      return _XmlHelper.domParserFactory().parseFromString(xmlString, "application/xml");
    }
    const anyGlobal = globalThis;
    const DP = anyGlobal?.DOMParser;
    if (typeof DP === "function") {
      return new DP().parseFromString(xmlString, "application/xml");
    }
    try {
      const xmldom = __require("@xmldom/xmldom");
      const Parser = xmldom.DOMParser || xmldom?.DOMParser;
      if (Parser) {
        return new Parser().parseFromString(xmlString, "application/xml");
      }
    } catch {
    }
    libWarn("No DOMParser available. Use XmlHelper.setDomParser() or install '@xmldom/xmldom'.");
    throw new Error("DOMParser not available in this environment");
  }
  /**
   * Gets a direct child element by local tag name
   */
  static getDirectChildrenByTagName(parent, tag) {
    return Array.from(parent.children).filter(
      (child) => child.localName === tag
    );
  }
  /**
   * Gets attribute value as number, defaulting to 0
   */
  static getAttrAsNumber(el, name) {
    const raw = el.getAttribute(name);
    if (raw == null || raw === "") return 0;
    const n = Number(raw);
    return Number.isFinite(n) ? n : 0;
  }
  static getColorFromElement(el, themeColors) {
    if (!el) return void 0;
    const srgb = el.getElementsByTagNameNS("*", "srgbClr")[0];
    if (srgb) {
      const val = srgb.getAttribute("val");
      if (!val) return void 0;
      return applyColorModifiers(`#${val}`, srgb);
    }
    const scheme = el.getElementsByTagNameNS("*", "schemeClr")[0];
    if (scheme) {
      const val = scheme.getAttribute("val");
      if (val && themeColors) {
        const aliasMap = {
          bg1: "lt1",
          bg2: "lt2",
          tx1: "dk1",
          tx2: "dk2"
        };
        const resolvedKey = aliasMap[val] || val;
        const baseColor = themeColors[resolvedKey];
        if (!baseColor) return void 0;
        return applyColorModifiers(baseColor, scheme);
      }
      return void 0;
    }
    const sys = el.getElementsByTagNameNS("*", "sysClr")[0];
    if (sys) {
      const lastClr = sys.getAttribute("lastClr");
      if (!lastClr) return void 0;
      return applyColorModifiers(`#${lastClr}`, sys);
    }
    return void 0;
  }
  static extractThemeColors(themeDoc) {
    if (!themeDoc) return {};
    const NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const themeColors = {};
    const clrScheme = themeDoc.getElementsByTagNameNS(NS, "clrScheme")[0];
    if (!clrScheme) return {};
    for (const node of Array.from(clrScheme.children)) {
      const name = node.localName;
      const srgbClr = node.getElementsByTagNameNS(NS, "srgbClr")[0];
      const sysClr = node.getElementsByTagNameNS(NS, "sysClr")[0];
      const hex = srgbClr?.getAttribute("val") ?? sysClr?.getAttribute("lastClr");
      if (hex) {
        themeColors[name] = `#${hex}`;
      }
    }
    return themeColors;
  }
  /**
   * Extracts table styles (fills and text colors per region) from theme XML.
   * Returns a map keyed by styleId (GUID or name), with region color maps.
   */
  static extractThemeTableStyles(themeDoc) {
    const styles = {};
    if (!themeDoc) return styles;
    const themeColors = _XmlHelper.extractThemeColors(themeDoc);
    const tblStyleLst = themeDoc.getElementsByTagNameNS("*", "tblStyleLst")[0] || null;
    if (!tblStyleLst) return styles;
    const tblStyles = Array.from(tblStyleLst.getElementsByTagNameNS("*", "tblStyle"));
    for (const ts of tblStyles) {
      const id = ts.getAttribute("styleId") || ts.getAttribute("name") || "";
      if (!id) continue;
      const fills = {};
      const fontColors = {};
      const prNodes = Array.from(ts.getElementsByTagNameNS("*", "tblStylePr"));
      for (const pr of prNodes) {
        const type = pr.getAttribute("type") || pr.getAttribute("val") || "";
        if (!type) continue;
        const tcStyle = pr.getElementsByTagNameNS("*", "tcStyle")[0] || null;
        const tcPr = tcStyle?.getElementsByTagNameNS("*", "tcPr")[0] || null;
        const solidCandidates = [
          tcPr?.getElementsByTagNameNS("*", "solidFill")[0] || null,
          tcStyle?.getElementsByTagNameNS("*", "solidFill")[0] || null,
          pr.getElementsByTagNameNS("*", "solidFill")[0] || null
        ];
        let fillColor;
        for (const cand of solidCandidates) {
          if (cand && !fillColor) fillColor = _XmlHelper.getColorFromElement(cand, themeColors);
        }
        if (!fillColor) {
          const fillRef = tcStyle?.getElementsByTagNameNS("*", "fillRef")[0] || pr.getElementsByTagNameNS("*", "fillRef")[0] || null;
          fillColor = _XmlHelper.getColorFromElement(fillRef, themeColors);
        }
        if (fillColor) fills[type] = fillColor;
        const txStyle = pr.getElementsByTagNameNS("*", "tcTxStyle")[0] || null;
        const txFillSolid = txStyle?.getElementsByTagNameNS("*", "solidFill")[0] || null;
        let textColor = _XmlHelper.getColorFromElement(txFillSolid, themeColors);
        if (!textColor) {
          const fontRef = txStyle?.getElementsByTagNameNS("*", "fontRef")[0] || null;
          textColor = _XmlHelper.getColorFromElement(fontRef, themeColors);
        }
        if (!textColor) {
          const anyScheme = txStyle?.getElementsByTagNameNS("*", "schemeClr")[0] || null;
          textColor = _XmlHelper.getColorFromElement(anyScheme, themeColors);
        }
        if (textColor) fontColors[type] = textColor;
      }
      styles[id] = { fills, fontColors };
    }
    return styles;
  }
  /**
   * Parses a <gradFill> element into a GradientFill object.
   * Returns null if the element is not a valid gradient.
   */
  static getGradientFromElement(gradFill, themeColors) {
    if (!gradFill || gradFill.localName !== "gradFill") return null;
    const gsLst = gradFill.getElementsByTagNameNS("*", "gsLst")[0];
    if (!gsLst) return null;
    const gsElements = Array.from(gsLst.getElementsByTagNameNS("*", "gs"));
    if (gsElements.length === 0) return null;
    const stops = [];
    for (const gs of gsElements) {
      const pos = gs.getAttribute("pos");
      const offset = pos ? Number(pos) / 1e3 : 0;
      const color = _XmlHelper.getColorFromElement(gs, themeColors);
      if (color) {
        stops.push({ offset, color });
      }
    }
    if (stops.length === 0) return null;
    stops.sort((a, b) => a.offset - b.offset);
    const linEl = gradFill.getElementsByTagNameNS("*", "lin")[0];
    const pathEl = gradFill.getElementsByTagNameNS("*", "path")[0];
    if (pathEl) {
      return { type: "gradient", gradientType: "radial", stops };
    }
    let angle = 180;
    if (linEl) {
      const ang = linEl.getAttribute("ang");
      if (ang) {
        const oomlDeg = Number(ang) / 6e4;
        angle = (oomlDeg + 90) % 360;
      }
    }
    return { type: "gradient", gradientType: "linear", stops, angle };
  }
  /** Allow host to provide a DOM parser (e.g., new (require('@xmldom/xmldom').DOMParser)()) */
  static setDomParser(factory) {
    _XmlHelper.domParserFactory = factory;
  }
  /** Relationship lookup: by Type suffix (avoids querySelector CSS) */
  static findRelationshipByTypeSuffix(doc, suffix) {
    const rels = doc.getElementsByTagName("Relationship");
    for (const el of Array.from(rels)) {
      const t = el.getAttribute("Type") || "";
      if (t.endsWith(suffix)) return el;
    }
    return null;
  }
  /** Relationship lookup: by Id */
  static findRelationshipById(doc, id) {
    const rels = doc.getElementsByTagName("Relationship");
    for (const el of Array.from(rels)) {
      if (el.getAttribute("Id") === id) return el;
    }
    return null;
  }
};

export {
  XmlHelper
};
//# sourceMappingURL=chunk-JB6TXILZ.js.map