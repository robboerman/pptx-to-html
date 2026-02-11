import { DiagramElement, DiagramShape, Fill } from "../models/SlideElement";
import { getSvgPathForShape } from "./shapePathMap";

const EMU_TO_PX = 9525;

function fillToSvgColor(fill: Fill, fallback = "transparent"): string {
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

function buildGradientDef(
  fill: Fill,
  id: string,
): string | null {
  if (fill.type !== "gradient") return null;
  const stops = fill.stops
    .map((s) => `<stop offset="${s.offset}%" stop-color="${s.color}" />`)
    .join("");
  if (fill.gradientType === "radial") {
    return `<radialGradient id="${id}" cx="50%" cy="50%" r="50%">${stops}</radialGradient>`;
  }
  const angle = fill.angle ?? 180;
  const rad = ((angle - 90) * Math.PI) / 180;
  const x1 = 50 - Math.cos(rad) * 50;
  const y1 = 50 - Math.sin(rad) * 50;
  const x2 = 50 + Math.cos(rad) * 50;
  const y2 = 50 + Math.sin(rad) * 50;
  return `<linearGradient id="${id}" x1="${x1}%" y1="${y1}%" x2="${x2}%" y2="${y2}%">${stops}</linearGradient>`;
}

/**
 * For shapes whose OOXML geometry depends on the aspect ratio (e.g. homePlate,
 * chevron), compute the polygon points in absolute pixel coordinates.
 * The arrow/notch depth = min(width, height) / 2 per OOXML defaults.
 * Returns null for shapes that don't need dynamic computation.
 */
function getDynamicPolygon(
  preset: string,
  x: number,
  y: number,
  w: number,
  h: number,
): string | null {
  const dx = Math.min(w, h) / 2;
  switch (preset) {
    case "homePlate": {
      // Flat left, pointed right
      const pts = [
        [x, y],
        [x + w - dx, y],
        [x + w, y + h / 2],
        [x + w - dx, y + h],
        [x, y + h],
      ];
      return pts.map(([px, py]) => `${px},${py}`).join(" ");
    }
    case "chevron": {
      // Notched left, pointed right
      const pts = [
        [x, y],
        [x + w - dx, y],
        [x + w, y + h / 2],
        [x + w - dx, y + h],
        [x, y + h],
        [x + dx, y + h / 2],
      ];
      return pts.map(([px, py]) => `${px},${py}`).join(" ");
    }
    case "rightArrow":
    case "arrow": {
      const pts = [
        [x, y + h * 0.25],
        [x + w - dx, y + h * 0.25],
        [x + w - dx, y],
        [x + w, y + h / 2],
        [x + w - dx, y + h],
        [x + w - dx, y + h * 0.75],
        [x, y + h * 0.75],
      ];
      return pts.map(([px, py]) => `${px},${py}`).join(" ");
    }
    default:
      return null;
  }
}

function renderShapeSvg(
  shape: DiagramShape,
  index: number,
): { svg: string; defs: string } {
  const x = shape.position.x / EMU_TO_PX;
  const y = shape.position.y / EMU_TO_PX;
  const w = shape.size.width / EMU_TO_PX;
  const h = shape.size.height / EMU_TO_PX;

  if (w <= 0 || h <= 0) return { svg: "", defs: "" };

  const fillColor = fillToSvgColor(shape.fill);
  const strokeColor = shape.stroke?.color ?? "none";
  const strokeWidth =
    shape.stroke?.width && Number.isFinite(shape.stroke.width)
      ? shape.stroke.width
      : 0;
  const rotation = shape.rotationDeg || 0;

  let defs = "";
  let fillAttr = `fill="${fillColor}"`;

  // Handle gradient fills with SVG defs
  if (shape.fill.type === "gradient") {
    const gradId = `dg-grad-${index}`;
    const gradDef = buildGradientDef(shape.fill, gradId);
    if (gradDef) {
      defs = gradDef;
      fillAttr = `fill="url(#${gradId})"`;
    }
  }

  const rotTransform = rotation
    ? ` transform="rotate(${rotation} ${x + w / 2} ${y + h / 2})"`
    : "";

  // Generate aspect-ratio-aware polygon for geometry types where
  // the arrow depth depends on the shape's height, not width.
  const dynamicPoints = getDynamicPolygon(shape.presetGeometry, x, y, w, h);
  if (dynamicPoints) {
    const shapeMarkup = `<polygon points="${dynamicPoints}" ${fillAttr} stroke="${strokeColor}" stroke-width="${strokeWidth}"${rotTransform} />`;
    return { svg: shapeMarkup, defs };
  }

  const raw = getSvgPathForShape(shape.presetGeometry);
  const [typeRaw, ...rest] = raw.trim().split(/\s+/);
  const type = typeRaw.toUpperCase().replace("_ARROW", "");
  const data = rest.join(" ");

  let shapeMarkup = "";

  switch (type) {
    case "PATH": {
      shapeMarkup = `<g${rotTransform}>
        <svg x="${x}" y="${y}" width="${w}" height="${h}" viewBox="0 0 100 100" overflow="visible">
          <path d="${data}" ${fillAttr} stroke="${strokeColor}" stroke-width="${strokeWidth}" />
        </svg>
      </g>`;
      break;
    }
    case "POLYGON":
    default: {
      const coords = data
        .split(/[\s,]+/)
        .map((v) => parseFloat(v))
        .filter((v) => !isNaN(v));

      const scaledPairs: string[] = [];
      for (let i = 0; i < coords.length; i += 2) {
        const px = x + (coords[i] / 100) * w;
        const py = y + (coords[i + 1] / 100) * h;
        scaledPairs.push(`${px},${py}`);
      }

      shapeMarkup = `<polygon points="${scaledPairs.join(" ")}" ${fillAttr} stroke="${strokeColor}" stroke-width="${strokeWidth}"${rotTransform} />`;
      break;
    }
  }

  return { svg: shapeMarkup, defs };
}

function renderTextOverlay(
  shape: DiagramShape,
  index: number,
): string {
  if (!shape.textContent) return "";

  const x = shape.position.x / EMU_TO_PX;
  const y = shape.position.y / EMU_TO_PX;
  const w = shape.size.width / EMU_TO_PX;
  const h = shape.size.height / EMU_TO_PX;

  if (w <= 0 || h <= 0) return "";

  const font = shape.textContent.font;
  const align = shape.textContent.align?.horizontal ?? "center";
  const vAlign = shape.textContent.align?.vertical ?? "middle";

  const justifyContent =
    vAlign === "top"
      ? "flex-start"
      : vAlign === "bottom"
        ? "flex-end"
        : "center";

  return `<foreignObject x="${x}" y="${y}" width="${w}" height="${h}">
    <div xmlns="http://www.w3.org/1999/xhtml" style="
      width: 100%;
      height: 100%;
      display: flex;
      flex-direction: column;
      justify-content: ${justifyContent};
      align-items: center;
      text-align: ${align};
      font-family: '${font.name}', Calibri, sans-serif;
      font-size: ${font.size}px;
      color: ${font.color};
      padding: 4px;
      box-sizing: border-box;
      overflow: hidden;
      line-height: 1.2;
    ">${shape.textContent.html}</div>
  </foreignObject>`;
}

export function renderDiagramElement(el: DiagramElement): string {
  const x = el.position.x / EMU_TO_PX;
  const y = el.position.y / EMU_TO_PX;
  const w = el.size.width / EMU_TO_PX;
  const h = el.size.height / EMU_TO_PX;

  if (w <= 0 || h <= 0) return "";

  const allDefs: string[] = [];
  const allShapes: string[] = [];
  const allTexts: string[] = [];

  for (let i = 0; i < el.shapes.length; i++) {
    const shape = el.shapes[i];
    const { svg, defs } = renderShapeSvg(shape, i);
    if (defs) allDefs.push(defs);
    if (svg) allShapes.push(svg);
    const textOverlay = renderTextOverlay(shape, i);
    if (textOverlay) allTexts.push(textOverlay);
  }

  const defsBlock =
    allDefs.length > 0 ? `<defs>${allDefs.join("\n")}</defs>` : "";

  // Apply 3D transform to container if any shape has rotation3D
  const first3D = el.shapes.find(s => s.rotation3D);
  let containerTransform = "";
  if (first3D?.rotation3D) {
    const r3d = first3D.rotation3D;
    const parts: string[] = [];
    if (r3d.perspective && r3d.perspective > 0) {
      const perspectiveVal = Math.round(45 / Math.tan((r3d.perspective * Math.PI) / 360));
      parts.push(`perspective(${perspectiveVal}px)`);
    }
    if (r3d.rotX) parts.push(`rotateX(${r3d.rotX}deg)`);
    if (r3d.rotY) parts.push(`rotateY(${r3d.rotY}deg)`);
    if (r3d.rotZ) parts.push(`rotateZ(${r3d.rotZ}deg)`);
    if (parts.length > 0) {
      containerTransform = `transform: ${parts.join(" ")}; transform-origin: center;`;
    }
  }

  return `<div style="position: absolute; left: ${x}px; top: ${y}px; width: ${w}px; height: ${h}px; overflow: visible; ${containerTransform}">
    <svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg" style="overflow: visible;">
      ${defsBlock}
      ${allShapes.join("\n")}
      ${allTexts.join("\n")}
    </svg>
  </div>`;
}
