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
      // Skip shapes nested inside group shapes — those are handled by extractGroups
      if (this.isInsideGroup(shape)) continue;

      const el = this.parseShapeElement(shape, themeColors);
      if (el) elements.push(el);
    }

    // Extract group shapes with proper coordinate transforms
    const groupElements = this.extractGroups(spTree, themeColors);
    elements.push(...groupElements);

    return elements;
  }

  /** Check if a shape element is nested inside a <grpSp> (group shape) */
  private static isInsideGroup(shape: Element): boolean {
    let parent = shape.parentElement;
    while (parent) {
      if (parent.localName === "grpSp") return true;
      if (parent.localName === "spTree") return false;
      parent = parent.parentElement;
    }
    return false;
  }

  /** Parse a single shape/connector element into a ShapeElement */
  private static parseShapeElement(
    shape: Element,
    themeColors: Record<string, string>,
    offsetX = 0,
    offsetY = 0,
  ): ShapeElement | null {
    const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
    const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];

    const x = (off ? XmlHelper.getAttrAsNumber(off, "x") : 0) + offsetX;
    const y = (off ? XmlHelper.getAttrAsNumber(off, "y") : 0) + offsetY;
    const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000;
    const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000;
    const rotAttr = xfrm?.getAttribute("rot");
    const rotationDeg = rotAttr ? Number(rotAttr) / 60000 : undefined;
    const flipH = xfrm?.getAttribute("flipH") === "1";
    const flipV = xfrm?.getAttribute("flipV") === "1";

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

    return {
      type: "shape",
      shapeType,
      position: { x, y },
      size: { width: cx, height: cy },
      fill,
      stroke,
      rotationDeg,
      flipH: flipH || undefined,
      flipV: flipV || undefined,
      cornerRadiusPct,
    };
  }

  /** Extract shapes from group shapes, applying group transforms */
  private static extractGroups(spTree: Element, themeColors: Record<string, string>): ShapeElement[] {
    const elements: ShapeElement[] = [];
    const groups = spTree.getElementsByTagNameNS("*", "grpSp");

    for (const grp of Array.from(groups)) {
      // Only process top-level groups (not nested)
      if (this.isInsideGroup(grp)) continue;

      const grpSpPr = grp.getElementsByTagNameNS("*", "grpSpPr")[0];
      if (!grpSpPr) continue;

      const grpXfrm = grpSpPr.getElementsByTagNameNS("*", "xfrm")[0];
      if (!grpXfrm) continue;

      const grpOff = grpXfrm.getElementsByTagNameNS("*", "off")[0];
      const grpExt = grpXfrm.getElementsByTagNameNS("*", "ext")[0];
      const chOff = grpXfrm.getElementsByTagNameNS("*", "chOff")[0];
      const chExt = grpXfrm.getElementsByTagNameNS("*", "chExt")[0];
      if (!grpOff || !grpExt || !chOff || !chExt) continue;

      const gx = XmlHelper.getAttrAsNumber(grpOff, "x");
      const gy = XmlHelper.getAttrAsNumber(grpOff, "y");
      const gw = XmlHelper.getAttrAsNumber(grpExt, "cx");
      const gh = XmlHelper.getAttrAsNumber(grpExt, "cy");
      const cx0 = XmlHelper.getAttrAsNumber(chOff, "x");
      const cy0 = XmlHelper.getAttrAsNumber(chOff, "y");
      const cw = XmlHelper.getAttrAsNumber(chExt, "cx");
      const ch = XmlHelper.getAttrAsNumber(chExt, "cy");

      const grpFlipV = grpXfrm.getAttribute("flipV") === "1";
      const grpFlipH = grpXfrm.getAttribute("flipH") === "1";

      const scaleX = cw > 0 ? gw / cw : 1;
      const scaleY = ch > 0 ? gh / ch : 1;

      // Process child shapes
      const childShapes = [
        ...Array.from(grp.getElementsByTagNameNS("*", "sp")),
        ...Array.from(grp.getElementsByTagNameNS("*", "cxnSp")),
      ];

      for (const child of childShapes) {
        const el = this.parseShapeElement(child, themeColors);
        if (!el) continue;

        // Map from child coordinate space to slide coordinate space
        const childRelX = (el.position.x - cx0) * scaleX;
        const childRelY = (el.position.y - cy0) * scaleY;
        const childW = el.size.width * scaleX;
        const childH = el.size.height * scaleY;

        let slideX = gx + childRelX;
        let slideY = gy + childRelY;

        if (grpFlipH) {
          slideX = gx + gw - childRelX - childW;
        }
        if (grpFlipV) {
          slideY = gy + gh - childRelY - childH;
        }

        el.position = { x: slideX, y: slideY };
        el.size = { width: childW, height: childH };

        // Group flips also flip the visual rendering of child shapes
        if (grpFlipV) {
          el.flipV = !(el.flipV ?? false) || undefined;
        }
        if (grpFlipH) {
          el.flipH = !(el.flipH ?? false) || undefined;
        }

        elements.push(el);
      }
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
