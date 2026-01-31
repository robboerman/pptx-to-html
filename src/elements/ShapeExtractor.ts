import { ShapeElement, Fill, StrokeStyle, ArrowHead } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";

/**
 * Responsible for extracting shape elements (including connectors) from a slide XML node.
 */
export class ShapeExtractor {
  /**
   * Extracts shape and connector elements from the <spTree> element of the slide.
   * @param spTree The <spTree> element.
   * @param themeColors Theme color mapping.
   * @returns List of ShapeElement extracted.
   */
  static extract(spTree: Element | null, themeColors: Record<string, string>): ShapeElement[] {
    if (!spTree) return [];

    const elements: ShapeElement[] = [];

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
      const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000;
      const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000;
      const rotAttr = xfrm?.getAttribute("rot");
      const rotationDeg = rotAttr ? Number(rotAttr) / 60000 : undefined;

      const prstGeom = shape.getElementsByTagNameNS("*", "prstGeom")[0];
      const shapeType = prstGeom?.getAttribute("prst") ?? "rect";

      // Extract corner radius adjustment for roundRect
      let cornerRadiusPct: number | undefined = undefined;
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

      const element: ShapeElement = {
        type: "shape",
        shapeType,
        position: { x, y },
        size: { width: cx, height: cy },
        fill,
        stroke,
        rotationDeg,
        cornerRadiusPct,
      };

      elements.push(element);
    }

    return elements;
  }

  /** Extract fill from shape properties, with theme/style fallback */
  private static extractFill(
    spPr: Element | undefined,
    shape: Element,
    themeColors: Record<string, string>,
  ): Fill {
    if (spPr) {
      // Check direct children of spPr only (not descendants inside <ln> etc.)
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

    // Fallback to style/fillRef (theme color)
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
  private static extractStroke(
    spPr: Element | undefined,
    themeColors: Record<string, string>,
  ): StrokeStyle | undefined {
    if (!spPr) return undefined;

    const ln = spPr.getElementsByTagNameNS("*", "ln")[0];
    if (!ln) return undefined;

    // Check for explicit <noFill> inside <ln>
    for (const child of Array.from(ln.children)) {
      if (child.localName === "noFill") return undefined;
    }

    const borderFill = ln.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
    const color = XmlHelper.getColorFromElement(borderFill, themeColors) ?? "transparent";

    // Line width (w) in EMUs → px
    const wAttr = ln.getAttribute("w");
    let width = 1;
    if (wAttr) {
      const w = Number(wAttr);
      if (!isNaN(w)) {
        width = w / 9525;
      }
    }

    // Arrowheads
    const headEndEl = ln.getElementsByTagNameNS("*", "headEnd")[0] ?? null;
    const tailEndEl = ln.getElementsByTagNameNS("*", "tailEnd")[0] ?? null;

    const headEnd: ArrowHead | undefined = headEndEl
      ? {
          type: headEndEl.getAttribute("type") || undefined,
          w: headEndEl.getAttribute("w") || undefined,
          len: headEndEl.getAttribute("len") || undefined,
        }
      : undefined;

    const tailEnd: ArrowHead | undefined = tailEndEl
      ? {
          type: tailEndEl.getAttribute("type") || undefined,
          w: tailEndEl.getAttribute("w") || undefined,
          len: tailEndEl.getAttribute("len") || undefined,
        }
      : undefined;

    // Dash style
    const prstDash = ln.getElementsByTagNameNS("*", "prstDash")[0];
    const dashStyle = prstDash?.getAttribute("val") || undefined;

    // Only return stroke if there's meaningful content
    if (color === "transparent" && !headEnd && !tailEnd && !dashStyle && width <= 0) {
      return undefined;
    }

    return { color, width, dashStyle, headEnd, tailEnd };
  }
}
