import { ShapeElement, Fill, ArrowHead } from "../models/SlideElement";
import { getSvgPathForShape } from "./shapePathMap";

/** Extract a CSS-usable color string from a Fill, or fallback */
function fillToCssColor(fill: Fill, fallback = "transparent"): string {
  switch (fill.type) {
    case "solid": return fill.color;
    case "gradient": return fill.stops.length > 0 ? fill.stops[0].color : fallback;
    case "none": return "transparent";
    default: return fallback;
  }
}

/** Generate CSS background property for a fill */
function fillToCssBackground(fill: Fill): string {
  switch (fill.type) {
    case "solid":
      return `background-color: ${fill.color};`;
    case "gradient": {
      if (fill.stops.length === 0) return "background-color: transparent;";
      if (fill.stops.length === 1) return `background-color: ${fill.stops[0].color};`;
      const stops = fill.stops.map(s => `${s.color} ${s.offset}%`).join(", ");
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

/** Map OOXML preset dash name to SVG stroke-dasharray */
function dashStyleToSvg(dashStyle?: string): string | undefined {
  if (!dashStyle || dashStyle === "solid") return undefined;
  const map: Record<string, string> = {
    dash: "8,4",
    dot: "2,4",
    dashDot: "8,4,2,4",
    lgDash: "16,4",
    lgDashDot: "16,4,2,4",
    lgDashDotDot: "16,4,2,4,2,4",
    sysDash: "4,2",
    sysDot: "1,2",
    sysDashDot: "4,2,1,2",
    sysDashDotDot: "4,2,1,2,1,2",
  };
  return map[dashStyle];
}

/** Resolve stroke color: prefer explicit, fallback to fill color, then default */
function resolveStrokeColor(el: ShapeElement): string {
  const strokeColor = el.stroke?.color;
  if (strokeColor && strokeColor !== "transparent") return strokeColor;
  const fc = fillToCssColor(el.fill);
  if (fc !== "transparent") return fc;
  return "#000";
}

/**
 * Renders a shape element as an absolutely positioned HTML or SVG element.
 * Supports all recognized PPTX shape types using SVG when necessary.
 */
export function renderShapeElement(el: ShapeElement, options: { scaleStrokes?: boolean } = {}): string {
    const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
    const x = nf(el.position?.x, 0) / 9525;
    const y = nf(el.position?.y, 0) / 9525;
    const width = nf(el.size?.width, 0) / 9525;
    const height = nf(el.size?.height, 0) / 9525;

    const rotation = el.rotationDeg && !isNaN(el.rotationDeg) ? el.rotationDeg : 0;
    const rotationStyle = rotation ? `transform: rotate(${rotation}deg); transform-origin: center;` : "";
    const strokeWidth = el.stroke?.width && Number.isFinite(el.stroke.width) && el.stroke.width > 0
      ? el.stroke.width
      : 1;
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

    // Basic HTML shapes
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
        const sw = el.stroke?.width && Number.isFinite(el.stroke.width) && el.stroke.width > 0
          ? el.stroke.width : 1;
        const lineColor = resolveStrokeColor(el);
        const dashAttr = dashStyleToSvg(el.stroke?.dashStyle);
        const dashStr = dashAttr ? ` stroke-dasharray="${dashAttr}"` : "";

        // Determine line endpoints within the SVG viewBox
        const svgW = Math.max(width, sw * 2);
        const svgH = Math.max(height, sw * 2);
        const isVertical = width < 1;
        const isHorizontal = height < 1;
        let x1: number, y1: number, x2: number, y2: number;
        if (isVertical) {
            x1 = svgW / 2; y1 = 0; x2 = svgW / 2; y2 = svgH;
        } else if (isHorizontal) {
            x1 = 0; y1 = svgH / 2; x2 = svgW; y2 = svgH / 2;
        } else {
            x1 = 0; y1 = 0; x2 = svgW; y2 = svgH;
        }

        // Build arrowhead markers
        const defs = buildMarkerDefs(el.stroke?.headEnd, el.stroke?.tailEnd, lineColor);
        const markerStartAttr = defs.startId ? ` marker-start="url(#${defs.startId})"` : "";
        const markerEndAttr = defs.endId ? ` marker-end="url(#${defs.endId})"` : "";

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
        let radius: number;
        if (adjVal !== undefined && adjVal >= 0) {
            const shorter = Math.min(width, height);
            radius = (adjVal / 100000) * shorter;
        } else {
            radius = Math.min(16, Math.min(width, height) * 0.1);
        }
        return `<div style="${style}
      ${bgCss}
      border: ${strokeWidth}px solid ${strokeColor};
      border-radius: ${radius}px;
      box-sizing: border-box;"></div>`;
    }

    // SVG-based shapes using prefixed definition
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
      el.stroke?.width && Number.isFinite(el.stroke.width) ? el.stroke.width : undefined,
      rotation,
      el.stroke?.headEnd,
      el.stroke?.tailEnd,
      options.scaleStrokes === true,
      dashStyleToSvg(el.stroke?.dashStyle)
    );
}

function shapeSvg(
  x: number,
  y: number,
  width: number,
  height: number,
  fill: string,
  stroke: string | undefined,
  raw: string,
  strokeWidthPx?: number,
  rotationDeg?: number,
  headEnd?: ArrowHead,
  tailEnd?: ArrowHead,
  scaleStrokes?: boolean,
  dashArray?: string
): string {
  const strokeColorOpt = stroke && stroke !== "transparent" ? stroke : undefined;
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
      const markerStartAttr = defs.startId ? `marker-start=\"url(#${defs.startId})\"` : "";
      const markerEndAttr = defs.endId ? `marker-end=\"url(#${defs.endId})\"` : "";
      return `<svg viewBox="0 0 100 100" style="${commonStyle}" overflow="visible">
        ${defs.defs}
        <path d="${data}" fill="none" stroke="${strokeColorOpt || "#000"}" stroke-width="${sw}" ${dashAttr} ${scaleStrokes ? "" : "vector-effect=\"non-scaling-stroke\""} ${markerStartAttr} ${markerEndAttr} />
      </svg>`;
    }

    case "POLYLINE":
    case "LINE": {
      const coords = data
        .split(/[\s,]+/)
        .map((v) => parseFloat(v))
        .filter((v) => !isNaN(v));

      if (coords.length < 4 || coords.length % 2 !== 0) {
        if (typeof console !== "undefined" && console.warn) {
          console.warn(`[pptx-to-html] Invalid POLYLINE/LINE shape data: "${data}"`);
        }
        return "";
      }

      const pointPairs: string[] = [];
      for (let i = 0; i < coords.length; i += 2) {
        pointPairs.push(`${coords[i]},${coords[i + 1]}`);
      }

      const effectiveWidth = width > 0 ? width : Math.max(sw * 2, 2);
      const effectiveHeight = height > 0 ? height : Math.max(sw * 2, 2);

      const defs = buildMarkerDefs(headEnd, tailEnd ?? (isArrow ? { type: "triangle", w: "med", len: "med" } : undefined), strokeColorOpt || "#000");
      const markerStartAttr = defs.startId ? `marker-start=\"url(#${defs.startId})\"` : "";
      const markerEndAttr = defs.endId ? `marker-end=\"url(#${defs.endId})\"` : "";

      const scaledPairs: string[] = [];
      for (let i = 0; i < coords.length; i += 2) {
        const px = (coords[i] / 100) * effectiveWidth;
        const py = (coords[i + 1] / 100) * effectiveHeight;
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
                    ${scaleStrokes ? "" : "vector-effect=\"non-scaling-stroke\""}
                    ${markerStartAttr} ${markerEndAttr} />
        </svg>`;
    }

    case "POLYGON":
    default:
      return `<svg viewBox="0 0 100 100" style="${commonStyle}">
        <polygon points="${data}" fill="${fill}" stroke="${strokeColorOpt ?? "none"}" stroke-width="${sw}" ${dashAttr} ${scaleStrokes ? "" : "vector-effect=\"non-scaling-stroke\""} />
      </svg>`;
  }
}

function buildMarkerDefs(
  headEnd: ArrowHead | undefined,
  tailEnd: ArrowHead | undefined,
  color: string
): { defs: string; startId?: string; endId?: string } {
  const parts: string[] = [];
  let startId: string | undefined;
  let endId: string | undefined;

  if (headEnd && headEnd.type && headEnd.type !== "none") {
    startId = `mstart-${Math.random().toString(36).slice(2, 8)}`;
    parts.push(markerDef(startId, headEnd, color));
  }
  if (tailEnd && tailEnd.type && tailEnd.type !== "none") {
    endId = `mend-${Math.random().toString(36).slice(2, 8)}`;
    parts.push(markerDef(endId, tailEnd, color));
  }

  return { defs: parts.length ? `<defs>${parts.join("\n")}</defs>` : "", startId, endId };
}

function markerDef(id: string, spec: ArrowHead, color: string): string {
  const sizeFactor = mapLen(spec.len);
  const base = 4 * sizeFactor;
  const refX = base;
  const refY = base / 2;

  switch ((spec.type || "triangle").toLowerCase()) {
    case "diamond":
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${base}" markerHeight="${base}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <polygon points="${base/2},0 ${base},${base/2} ${base/2},${base} 0,${base/2}" fill="${color}" />
              </marker>`;
    case "oval":
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${base}" markerHeight="${base}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <circle cx="${base/2}" cy="${base/2}" r="${base/2}" fill="${color}" />
              </marker>`;
    case "stealth":
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${base}" markerHeight="${base}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <polygon points="${base},${base/2} 0,0 0,${base}" fill="${color}" />
              </marker>`;
    case "arrow":
    case "triangle":
    default:
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${base}" markerHeight="${base}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <polygon points="0,0 ${base},${base/2} 0,${base}" fill="${color}" />
              </marker>`;
  }
}

function mapLen(len?: string): number {
  switch ((len || "med").toLowerCase()) {
    case "sm":
    case "small":
      return 1.5;
    case "lg":
    case "large":
      return 2.5;
    case "med":
    case "medium":
    default:
      return 2;
  }
}
