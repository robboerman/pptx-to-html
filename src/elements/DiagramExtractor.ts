import JSZip from "jszip";
import { XmlHelper } from "../core/XmlHelper";
import {
  DiagramElement,
  DiagramShape,
  Fill,
  StrokeStyle,
  ShapeTextContent,
  computeZOrder,
} from "../models/SlideElement";

const DIAGRAM_URI =
  "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DIAGRAM_DRAWING_REL_TYPE =
  "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";

export class DiagramExtractor {
  static async extract(
    spTree: Element | null,
    relsXml: Document,
    zip: JSZip,
    themeColors: Record<string, string>,
  ): Promise<DiagramElement[]> {
    if (!spTree) return [];

    const diagrams: DiagramElement[] = [];

    // Collect all diagramDrawing relationships from rels
    const drawingRels: Element[] = [];
    const allRels = relsXml.getElementsByTagName("Relationship");
    for (const rel of Array.from(allRels)) {
      const type = rel.getAttribute("Type") || "";
      if (type === DIAGRAM_DRAWING_REL_TYPE) {
        drawingRels.push(rel);
      }
    }

    if (drawingRels.length === 0) return [];

    const gFrames = spTree.getElementsByTagNameNS("*", "graphicFrame");
    let drawingRelIndex = 0;

    for (const gf of Array.from(gFrames)) {
      const graphicData =
        gf.getElementsByTagNameNS("*", "graphicData")[0] ?? null;
      if (!graphicData) continue;

      const uri = graphicData.getAttribute("uri");
      if (uri !== DIAGRAM_URI) continue;

      // Get graphicFrame position/size
      const xfrm = gf.getElementsByTagNameNS("*", "xfrm")[0] ?? null;
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;
      const frameX = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
      const frameY = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;
      const frameCx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000;
      const frameCy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 600000;

      // Match this diagram to its drawing relationship
      if (drawingRelIndex >= drawingRels.length) break;
      const drawingRel = drawingRels[drawingRelIndex];
      drawingRelIndex++;

      const target = drawingRel.getAttribute("Target") || "";
      if (!target) continue;

      const fullPath = this.resolvePath(target, "ppt/slides");
      const file = zip.file(fullPath);
      if (!file) continue;

      const xmlStr = await file.async("string");
      const drawingDoc = XmlHelper.parseXml(xmlStr);

      const shapes = this.parseDrawingShapes(
        drawingDoc,
        frameCx,
        frameCy,
        themeColors,
      );

      if (shapes.length === 0) continue;

      diagrams.push({
        type: "diagram",
        position: { x: frameX, y: frameY },
        size: { width: frameCx, height: frameCy },
        shapes,
        zOrder: computeZOrder(gf),
      });
    }

    return diagrams;
  }

  private static parseDrawingShapes(
    drawingDoc: Document,
    _frameCx: number,
    _frameCy: number,
    themeColors: Record<string, string>,
  ): DiagramShape[] {
    const shapes: DiagramShape[] = [];

    // Find all <dsp:sp> elements in the drawing's spTree
    const spElements = drawingDoc.getElementsByTagNameNS("*", "sp");
    if (spElements.length === 0) return [];

    for (const sp of Array.from(spElements)) {
      const spPr = sp.getElementsByTagNameNS("*", "spPr")[0] ?? null;
      if (!spPr) continue;

      const xfrm = spPr.getElementsByTagNameNS("*", "xfrm")[0] ?? null;
      if (!xfrm) continue;

      const offEl = xfrm.getElementsByTagNameNS("*", "off")[0] ?? null;
      const extEl = xfrm.getElementsByTagNameNS("*", "ext")[0] ?? null;
      if (!offEl || !extEl) continue;

      const x = XmlHelper.getAttrAsNumber(offEl, "x");
      const y = XmlHelper.getAttrAsNumber(offEl, "y");
      const cx = XmlHelper.getAttrAsNumber(extEl, "cx");
      const cy = XmlHelper.getAttrAsNumber(extEl, "cy");

      if (cx === 0 && cy === 0) continue;

      // Preset geometry
      const prstGeom =
        spPr.getElementsByTagNameNS("*", "prstGeom")[0] ?? null;
      const presetGeometry = prstGeom?.getAttribute("prst") ?? "rect";

      // Rotation
      const rotAttr = xfrm.getAttribute("rot");
      const rotationDeg = rotAttr ? Number(rotAttr) / 60000 : undefined;

      // 3D Rotation
      const rotation3D = XmlHelper.get3DRotation(spPr);

      // Fill
      const fill = this.extractFill(spPr, sp, themeColors);

      // Stroke
      const stroke = this.extractStroke(spPr, themeColors);

      // Text content
      const textContent = this.extractTextContent(sp, themeColors);

      // Drawing coordinates are already in the frame's EMU coordinate space
      shapes.push({
        presetGeometry,
        position: { x, y },
        size: { width: cx, height: cy },
        fill,
        stroke,
        rotationDeg,
        rotation3D,
        textContent,
      });
    }

    return shapes;
  }

  private static extractFill(
    spPr: Element,
    sp: Element,
    themeColors: Record<string, string>,
  ): Fill {
    // Check direct children of spPr
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

    // Fallback to style/fillRef
    const style = sp.getElementsByTagNameNS("*", "style")[0];
    const fillRef = style?.getElementsByTagNameNS("*", "fillRef")[0];
    const schemeClr = fillRef?.getElementsByTagNameNS("*", "schemeClr")[0];
    const val = schemeClr?.getAttribute("val");
    if (val && themeColors[val]) {
      return { type: "solid", color: themeColors[val] };
    }

    return { type: "none" };
  }

  private static extractStroke(
    spPr: Element,
    themeColors: Record<string, string>,
  ): StrokeStyle | undefined {
    const ln = spPr.getElementsByTagNameNS("*", "ln")[0];
    if (!ln) return undefined;

    // Check for <noFill> inside <ln>
    for (const child of Array.from(ln.children)) {
      if (child.localName === "noFill") return undefined;
    }

    const borderFill =
      ln.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
    const color =
      XmlHelper.getColorFromElement(borderFill, themeColors) ?? "transparent";

    const wAttr = ln.getAttribute("w");
    let width = 1;
    if (wAttr) {
      const w = Number(wAttr);
      if (!isNaN(w)) {
        width = w / 9525;
      }
    }

    const prstDash = ln.getElementsByTagNameNS("*", "prstDash")[0];
    const dashStyle = prstDash?.getAttribute("val") || undefined;

    if (
      color === "transparent" &&
      !dashStyle &&
      width <= 0
    ) {
      return undefined;
    }

    return { color, width, dashStyle };
  }

  private static extractTextContent(
    sp: Element,
    themeColors: Record<string, string>,
  ): ShapeTextContent | undefined {
    // Try dsp:txBody or just txBody
    const txBody = sp.getElementsByTagNameNS("*", "txBody")[0] ?? null;
    if (!txBody) return undefined;

    const paragraphs = txBody.getElementsByTagNameNS("*", "p");
    if (paragraphs.length === 0) return undefined;

    let allText = "";
    let fontName = "Calibri";
    let fontSize = 12;
    let fontColor: string | undefined;
    let hAlign: "left" | "center" | "right" | "justify" | undefined;

    const htmlParts: string[] = [];

    for (const para of Array.from(paragraphs)) {
      const pPr = para.getElementsByTagNameNS("*", "pPr")[0] ?? null;
      const algn = pPr?.getAttribute("algn");
      if (algn === "ctr") hAlign = "center";
      else if (algn === "r") hAlign = "right";
      else if (algn === "just") hAlign = "justify";
      else if (algn === "l") hAlign = "left";

      const runs = para.getElementsByTagNameNS("*", "r");
      const paraTexts: string[] = [];

      for (const run of Array.from(runs)) {
        const rPr = run.getElementsByTagNameNS("*", "rPr")[0] ?? null;
        const t = run.getElementsByTagNameNS("*", "t")[0]?.textContent || "";
        if (!t) continue;

        paraTexts.push(t);
        allText += t;

        if (rPr) {
          const sz = rPr.getAttribute("sz");
          if (sz) {
            const n = parseInt(sz, 10);
            if (Number.isFinite(n)) fontSize = n / 100;
          }

          const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
          const typeface = latin?.getAttribute("typeface");
          if (typeface) fontName = typeface;

          const solidFill =
            rPr.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
          const color = XmlHelper.getColorFromElement(
            solidFill,
            themeColors,
          );
          if (color) fontColor = color;
        }
      }

      if (paraTexts.length > 0) {
        htmlParts.push(paraTexts.map(t => this.escapeHtml(t)).join(""));
      }
    }

    if (!allText.trim()) return undefined;

    // Resolve text color from shape style fontRef if not set explicitly
    if (!fontColor) {
      const style = sp.getElementsByTagNameNS("*", "style")[0] ?? null;
      const fontRef = style?.getElementsByTagNameNS("*", "fontRef")[0] ?? null;
      if (fontRef) {
        const color = XmlHelper.getColorFromElement(fontRef, themeColors);
        if (color) fontColor = color;
      }
    }

    const html = htmlParts.join("<br/>");

    return {
      html,
      font: { name: fontName, size: fontSize, color: fontColor ?? "#000000" },
      align: hAlign ? { horizontal: hAlign, vertical: "middle" } : { vertical: "middle" },
    };
  }

  private static escapeHtml(text: string): string {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  private static resolvePath(target: string, baseDir: string): string {
    const parts = (baseDir + "/" + target).split("/");
    const resolved: string[] = [];
    for (const part of parts) {
      if (part === "..") resolved.pop();
      else if (part !== ".") resolved.push(part);
    }
    return resolved.join("/");
  }
}
