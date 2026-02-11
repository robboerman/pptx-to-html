// Lightweight logging helper to centralize library warnings
function libWarn(msg: string) {
  if (typeof console !== "undefined" && console.warn) {
    console.warn(`[pptx-to-html] ${msg}`);
  }
}

type DomParserLike = { parseFromString(xml: string, mimeType: string): Document };

/** Parse a "#RRGGBB" hex string into [r, g, b] (0-255) */
function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  return [
    parseInt(h.substring(0, 2), 16),
    parseInt(h.substring(2, 4), 16),
    parseInt(h.substring(4, 6), 16),
  ];
}

/** Convert [r, g, b] (0-255) to "#RRGGBB" */
function rgbToHex(r: number, g: number, b: number): string {
  const clamp = (v: number) => Math.max(0, Math.min(255, Math.round(v)));
  return `#${[clamp(r), clamp(g), clamp(b)].map(v => v.toString(16).padStart(2, "0")).join("")}`;
}

/** Convert [r, g, b] (0-255) to [h (0-360), s (0-1), l (0-1)] */
function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
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

/** Convert [h (0-360), s (0-1), l (0-1)] to [r, g, b] (0-255) */
function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) {
    const v = Math.round(l * 255);
    return [v, v, v];
  }
  const hue2rgb = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hn = h / 360;
  return [
    Math.round(hue2rgb(p, q, hn + 1 / 3) * 255),
    Math.round(hue2rgb(p, q, hn) * 255),
    Math.round(hue2rgb(p, q, hn - 1 / 3) * 255),
  ];
}

/**
 * Apply OOXML color modifiers (lumMod, lumOff, tint, shade, satMod, satOff)
 * found as child elements of a color element (srgbClr, schemeClr, sysClr).
 */
function applyColorModifiers(hex: string, colorEl: Element): string {
  const children = Array.from(colorEl.children);
  if (children.length === 0) return hex;

  let [r, g, b] = hexToRgb(hex);

  // Collect modifier values (val is 0–100000, representing percentage * 1000)
  let tint: number | undefined;
  let shade: number | undefined;
  let lumMod: number | undefined;
  let lumOff: number | undefined;
  let satMod: number | undefined;
  let satOff: number | undefined;
  let alpha: number | undefined;

  for (const child of children) {
    const tag = child.localName;
    const val = Number(child.getAttribute("val") || "0");
    if (!Number.isFinite(val)) continue;
    switch (tag) {
      case "tint": tint = val; break;
      case "shade": shade = val; break;
      case "lumMod": lumMod = val; break;
      case "lumOff": lumOff = val; break;
      case "satMod": satMod = val; break;
      case "satOff": satOff = val; break;
      case "alpha": alpha = val; break;
    }
  }

  // Apply tint (lighten towards white) — RGB space
  if (tint !== undefined) {
    const t = tint / 100000;
    r = Math.round(r + (255 - r) * t);
    g = Math.round(g + (255 - g) * t);
    b = Math.round(b + (255 - b) * t);
  }

  // Apply shade (darken towards black) — RGB space
  if (shade !== undefined) {
    const s = shade / 100000;
    r = Math.round(r * s);
    g = Math.round(g * s);
    b = Math.round(b * s);
  }

  // Apply HSL-based modifiers
  if (lumMod !== undefined || lumOff !== undefined || satMod !== undefined || satOff !== undefined) {
    let [h, s, l] = rgbToHsl(r, g, b);
    if (lumMod !== undefined) l = l * (lumMod / 100000);
    if (lumOff !== undefined) l = l + (lumOff / 100000);
    if (satMod !== undefined) s = s * (satMod / 100000);
    if (satOff !== undefined) s = s + (satOff / 100000);
    l = Math.max(0, Math.min(1, l));
    s = Math.max(0, Math.min(1, s));
    [r, g, b] = hslToRgb(h, s, l);
  }

  if (alpha !== undefined) {
    const a = Math.max(0, Math.min(1, alpha / 100000));
    return `rgba(${r},${g},${b},${a})`;
  }
  return rgbToHex(r, g, b);
}

export class XmlHelper {
  private static domParserFactory: (() => DomParserLike) | null = null;
  /**
   * Parses a string containing XML into a DOM Document
   * @param xmlString XML string to parse
   * @returns DOM Document
   */
  static parseXml(xmlString: string): Document {
    if (XmlHelper.domParserFactory) {
      return XmlHelper.domParserFactory().parseFromString(xmlString, "application/xml");
    }
    const anyGlobal = globalThis as any;
    const DP: any = anyGlobal?.DOMParser;
    if (typeof DP === "function") {
      return new DP().parseFromString(xmlString, "application/xml");
    }
    try {
      // Optional runtime load if host app installed it; not a hard dependency
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      const xmldom = require("@xmldom/xmldom");
      const Parser = xmldom.DOMParser || xmldom?.DOMParser;
      if (Parser) {
        return new Parser().parseFromString(xmlString, "application/xml");
      }
    } catch {
      // ignore
    }
    libWarn("No DOMParser available. Use XmlHelper.setDomParser() or install '@xmldom/xmldom'.");
    throw new Error("DOMParser not available in this environment");
  }

  /**
   * Gets a direct child element by local tag name
   */
  static getDirectChildrenByTagName(
    parent: Element,
    tag: string
  ): Element[] {
    return Array.from(parent.children).filter(
      (child) => child.localName === tag
    );
  }

  /**
   * Gets attribute value as number, defaulting to 0
   */
  static getAttrAsNumber(el: Element, name: string): number {
    const raw = el.getAttribute(name);
    if (raw == null || raw === "") return 0;
    const n = Number(raw);
    return Number.isFinite(n) ? n : 0;
  }

  static getColorFromElement(el: Element | null, themeColors?: Record<string, string>): string | undefined {
    if (!el) return undefined;

    // 1. Try <srgbClr val="...">
    const srgb = el.getElementsByTagNameNS("*", "srgbClr")[0];
    if (srgb) {
      const val = srgb.getAttribute("val");
      if (!val) return undefined;
      return applyColorModifiers(`#${val}`, srgb);
    }

    // 2. Try <schemeClr val="..."> resolved via themeColors (including aliases)
    const scheme = el.getElementsByTagNameNS("*", "schemeClr")[0];
    if (scheme) {
      const val = scheme.getAttribute("val");
      if (val && themeColors) {
        const aliasMap: Record<string, string> = {
          bg1: "lt1",
          bg2: "lt2",
          tx1: "dk1",
          tx2: "dk2"
        };
        const resolvedKey = aliasMap[val] || val;
        const baseColor = themeColors[resolvedKey];
        if (!baseColor) return undefined;
        return applyColorModifiers(baseColor, scheme);
      }
      return undefined;
    }

    // 3. Try <sysClr lastClr="...">
    const sys = el.getElementsByTagNameNS("*", "sysClr")[0];
    if (sys) {
      const lastClr = sys.getAttribute("lastClr");
      if (!lastClr) return undefined;
      return applyColorModifiers(`#${lastClr}`, sys);
    }

    return undefined;
  }

  static extractThemeColors(themeDoc: Document | null): Record<string, string> {
    if (!themeDoc) return {};

    const NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const themeColors: Record<string, string> = {};

    const clrScheme = themeDoc.getElementsByTagNameNS(NS, "clrScheme")[0];
    if (!clrScheme) return {};

    for (const node of Array.from(clrScheme.children)) {
      const name = node.localName; // e.g., bg1, tx1, accent1...

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
  static extractThemeTableStyles(themeDoc: Document | null): Record<string, { fills: Record<string, string>; fontColors: Record<string, string> }> {
    const styles: Record<string, { fills: Record<string, string>; fontColors: Record<string, string> }> = {};
    if (!themeDoc) return styles;

    const themeColors = XmlHelper.extractThemeColors(themeDoc);
    const tblStyleLst = themeDoc.getElementsByTagNameNS("*", "tblStyleLst")[0] || null;
    if (!tblStyleLst) return styles;

    const tblStyles = Array.from(tblStyleLst.getElementsByTagNameNS("*", "tblStyle"));
    for (const ts of tblStyles) {
      const id = ts.getAttribute("styleId") || ts.getAttribute("name") || "";
      if (!id) continue;
      const fills: Record<string, string> = {};
      const fontColors: Record<string, string> = {};

      const prNodes = Array.from(ts.getElementsByTagNameNS("*", "tblStylePr"));
      for (const pr of prNodes) {
        const type = pr.getAttribute("type") || pr.getAttribute("val") || ""; // wholeTbl, firstRow, band1H, band2H, band1V, band2V, firstCol, lastCol, lastRow
        if (!type) continue;

        // Resolve fill color: try tcStyle/tcPr/solidFill, then any solidFill under tcStyle, then fillRef, then direct solidFill
        const tcStyle = pr.getElementsByTagNameNS("*", "tcStyle")[0] || null;
        const tcPr = tcStyle?.getElementsByTagNameNS("*", "tcPr")[0] || null;
        const solidCandidates: (Element | null)[] = [
          tcPr?.getElementsByTagNameNS("*", "solidFill")[0] || null,
          tcStyle?.getElementsByTagNameNS("*", "solidFill")[0] || null,
          pr.getElementsByTagNameNS("*", "solidFill")[0] || null,
        ];
        let fillColor: string | undefined;
        for (const cand of solidCandidates) {
          if (cand && !fillColor) fillColor = XmlHelper.getColorFromElement(cand, themeColors);
        }
        if (!fillColor) {
          const fillRef = tcStyle?.getElementsByTagNameNS("*", "fillRef")[0] || pr.getElementsByTagNameNS("*", "fillRef")[0] || null;
          fillColor = XmlHelper.getColorFromElement(fillRef, themeColors);
        }
        if (fillColor) fills[type] = fillColor;

        // Resolve text color: try tcTxStyle/txFill/solidFill, any solidFill, then fontRef (schemeClr)
        const txStyle = pr.getElementsByTagNameNS("*", "tcTxStyle")[0] || null;
        const txFillSolid = txStyle?.getElementsByTagNameNS("*", "solidFill")[0] || null;
        let textColor = XmlHelper.getColorFromElement(txFillSolid, themeColors);
        if (!textColor) {
          const fontRef = txStyle?.getElementsByTagNameNS("*", "fontRef")[0] || null;
          textColor = XmlHelper.getColorFromElement(fontRef as any, themeColors);
        }
        if (!textColor) {
          const anyScheme = txStyle?.getElementsByTagNameNS("*", "schemeClr")[0] || null;
          textColor = XmlHelper.getColorFromElement(anyScheme as any, themeColors);
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
  static getGradientFromElement(
    gradFill: Element | null,
    themeColors?: Record<string, string>
  ): import("../models/SlideElement").GradientFill | null {
    if (!gradFill || gradFill.localName !== "gradFill") return null;

    const gsLst = gradFill.getElementsByTagNameNS("*", "gsLst")[0];
    if (!gsLst) return null;

    const gsElements = Array.from(gsLst.getElementsByTagNameNS("*", "gs"));
    if (gsElements.length === 0) return null;

    const stops: { offset: number; color: string }[] = [];
    for (const gs of gsElements) {
      const pos = gs.getAttribute("pos");
      const offset = pos ? Number(pos) / 1000 : 0; // pos is 0–100000 → 0–100
      const color = XmlHelper.getColorFromElement(gs, themeColors);
      if (color) {
        stops.push({ offset, color });
      }
    }

    if (stops.length === 0) return null;
    stops.sort((a, b) => a.offset - b.offset);

    // Determine gradient type and angle
    const linEl = gradFill.getElementsByTagNameNS("*", "lin")[0];
    const pathEl = gradFill.getElementsByTagNameNS("*", "path")[0];

    if (pathEl) {
      // Radial gradient
      return { type: "gradient", gradientType: "radial", stops };
    }

    // Linear gradient (default)
    let angle = 180; // default: top-to-bottom in CSS
    if (linEl) {
      const ang = linEl.getAttribute("ang");
      if (ang) {
        // OOXML: ang in 60000ths of a degree, 0 = left-to-right
        // CSS: 0deg = bottom-to-top, 90deg = left-to-right
        const oomlDeg = Number(ang) / 60000;
        angle = (oomlDeg + 90) % 360;
      }
    }

    return { type: "gradient", gradientType: "linear", stops, angle };
  }

  static get3DRotation(spPr: Element | null | undefined): import("../models/SlideElement").Rotation3D | undefined {
    if (!spPr) return undefined;
    const scene3d = spPr.getElementsByTagNameNS("*", "scene3d")[0];
    if (!scene3d) return undefined;
    const camera = scene3d.getElementsByTagNameNS("*", "camera")[0];
    if (!camera) return undefined;
    const rot = camera.getElementsByTagNameNS("*", "rot")[0];
    if (!rot) return undefined;

    const lat = rot.getAttribute("lat");
    const lon = rot.getAttribute("lon");
    const rev = rot.getAttribute("rev");

    // OOXML camera sphere: lat change = look up/down = CSS rotateX,
    // lon change = orbit horizontally = CSS rotateY
    const rotX = lat ? Number(lat) / 60000 : undefined;
    const rotY = lon ? Number(lon) / 60000 : undefined;
    const rotZ = rev ? Number(rev) / 60000 : undefined;

    if (rotX === undefined && rotY === undefined && rotZ === undefined) return undefined;

    const prst = camera.getAttribute("prst") || "";
    const isPerspectiveCamera = prst.startsWith("perspective");

    let perspective: number | undefined = undefined;
    if (isPerspectiveCamera) {
      const fov = camera.getAttribute("fov");
      const fovDeg = fov ? Number(fov) / 60000 : 45;
      // fov=0 means no perspective distortion even for perspective presets
      if (fovDeg > 0) {
        perspective = fovDeg;
      }
    }

    return { rotX, rotY, rotZ, perspective };
  }

  /** Allow host to provide a DOM parser (e.g., new (require('@xmldom/xmldom').DOMParser)()) */
  static setDomParser(factory: () => DomParserLike) {
    XmlHelper.domParserFactory = factory;
  }

  /** Relationship lookup: by Type suffix (avoids querySelector CSS) */
  static findRelationshipByTypeSuffix(doc: Document, suffix: string): Element | null {
    const rels = doc.getElementsByTagName("Relationship");
    for (const el of Array.from(rels)) {
      const t = el.getAttribute("Type") || "";
      if (t.endsWith(suffix)) return el;
    }
    return null;
  }

  /** Relationship lookup: by Id */
  static findRelationshipById(doc: Document, id: string): Element | null {
    const rels = doc.getElementsByTagName("Relationship");
    for (const el of Array.from(rels)) {
      if (el.getAttribute("Id") === id) return el;
    }
    return null;
  }
}
