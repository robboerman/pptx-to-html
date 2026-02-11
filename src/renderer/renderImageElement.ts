import { ImageElement } from "../models/SlideElement";

/**
 * Renders an image element as an absolutely positioned <img> tag.
 * @param el Image element to render.
 * @returns HTML string representing the image element.
 */
export function renderImageElement(el: ImageElement): string {
  const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
  const w = nf(el.size?.width, 0) / 9525;
  const h = nf(el.size?.height, 0) / 9525;

  let borderCss = "";
  if (el.stroke && el.stroke.color !== "transparent") {
    const bw = el.stroke.width > 0 ? el.stroke.width : 1;
    borderCss = `border: ${bw}px solid ${el.stroke.color};`;
  }

  const radiusCss = el.borderRadius ? `border-radius: ${el.borderRadius}px;` : "";

  return `<img src="${el.src}" style="
    position: absolute;
    left: ${nf(el.position?.x, 0) / 9525}px;
    top: ${nf(el.position?.y, 0) / 9525}px;
    width: ${w}px;
    height: ${h}px;
    object-fit: cover;
    box-sizing: border-box;
    ${borderCss}
    ${radiusCss}
  " />`;
}
