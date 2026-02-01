import { SlideElement, BackgroundElement, Fill } from "../models/SlideElement";
import { renderTextElement } from "./renderTextElement";
import { renderImageElement } from "./renderImageElement";
import { renderShapeElement } from "./renderShapeElement";
import { renderTableElement } from "./renderTableElement";
import { renderChartElement } from "./renderChartElement";
import { renderDiagramElement } from "./renderDiagramElement";

function backgroundFillToCss(fill: Fill): string {
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
      return `background-image: url('${fill.src}'); background-size: cover; background-position: center; background-repeat: no-repeat;`;
    case "none":
      return "background-color: transparent;";
  }
}

/**
 * Converts a list of SlideElements into an HTML string with absolute positioning.
 */
export class HtmlRenderer {
  /**
   * Renders a slide to an HTML <div> with all elements positioned accordingly.
   * @param elements List of SlideElement (text, image, shape)
   * @param options Optional width and height (in px) for the slide container.
   *  - If scaleToFit=true, width/height define the outer container size and contents are scaled from the base 960x540.
   * @returns HTML string representing the slide
   */
  static render(
    elements: SlideElement[],
    options: { width?: number; height?: number; scaleToFit?: boolean; letterbox?: boolean; baseWidth?: number; baseHeight?: number } = {}
  ): string {
    const baseW = options.baseWidth ?? 960;
    const baseH = options.baseHeight ?? 540;
    const targetW = options.width ?? baseW;
    const targetH = options.height ?? baseH;
    const scaleToFit = options.scaleToFit === true;
    // Default letterbox to true when scaleToFit is enabled unless explicitly set to false
    const letterbox = scaleToFit ? options.letterbox !== false : options.letterbox === true;

    const htmlParts = elements.map((el) => {
      switch (el.type) {
        case "background": {
          const bgEl = el as BackgroundElement;
          const styleBg = backgroundFillToCss(bgEl.fill);
          return `<div style="position:absolute; left:0; top:0; width:${baseW}px; height:${baseH}px; ${styleBg}"></div>`;
        }
        case "text": return renderTextElement(el);
        case "image": return renderImageElement(el);
        case "shape": return renderShapeElement(el, { scaleStrokes: scaleToFit });
        case "table": return renderTableElement(el);
        case "chart": return renderChartElement(el as any);
        case "diagram": return renderDiagramElement(el as any);
        default:
          if (typeof console !== "undefined" && console.warn) {
            console.warn(`[pptx-to-html] Unsupported element type: ${(el as any)?.type}`);
          }
          return "";
      }
    });

    if (scaleToFit) {
      if (letterbox) {
        const s = Math.min(targetW / baseW, targetH / baseH);
        const offsetX = (targetW - baseW * s) / 2;
        const offsetY = (targetH - baseH * s) / 2;
        return (
          `<div class="slide-container" style="position: relative; width: ${targetW}px; height: ${targetH}px; overflow: hidden; background-color: #000;">
            <div class="slide" style="position: absolute; left: ${offsetX}px; top: ${offsetY}px; width: ${baseW}px; height: ${baseH}px; transform: scale(${s}); transform-origin: top left; background-color: #fff;">
              ${htmlParts.join("\n")}
            </div>
          </div>`
        );
      } else {
        const sx = targetW / baseW;
        const sy = targetH / baseH;
        return (
          `<div class="slide-container" style="position: relative; width: ${targetW}px; height: ${targetH}px; overflow: hidden;">
            <div class="slide" style="position: absolute; left: 0; top: 0; width: ${baseW}px; height: ${baseH}px; transform: scale(${sx}, ${sy}); transform-origin: top left; background-color: #fff;">
              ${htmlParts.join("\n")}
            </div>
          </div>`
        );
      }
    }

    return (
      `<div class="slide" style="position: relative; width: ${targetW}px; height: ${targetH}px; overflow: hidden; background-color: #fff;">
        ${htmlParts.join("\n")}
      </div>`
    );
  }
}
