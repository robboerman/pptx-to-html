import { ImageElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";
import JSZip from "jszip";

/**
 * Responsible for extracting image elements from a slide XML node.
 */
export class ImageExtractor {
  /**
   * Extracts image elements from the <spTree> element using rels from slide relationships.
   * @param spTree The <spTree> element of the slide.
   * @param rels XML Document for slide relationships (ppt/slides/_rels/slideX.xml.rels).
   * @param zip The JSZip archive of the entire .pptx file.
   * @returns List of ImageElement extracted.
   */
  static async extract(
    spTree: Element | null,
    rels: Document,
    zip: JSZip,
    basePath: string = "ppt/slides"
  ): Promise<ImageElement[]> {
    if (!spTree) return [];

    const elements: ImageElement[] = [];

    const pics = spTree.getElementsByTagNameNS("*", "pic");
    for (const pic of Array.from(pics)) {
      const blip = pic.getElementsByTagNameNS("*", "blip")[0];
      const embedId = blip?.getAttribute("r:embed") ?? "";
      if (!embedId) continue;

      const relEl = (rels && (rels as any).getElementsByTagName) ? (function(){
        const els = rels.getElementsByTagName("Relationship");
        for (const e of Array.from(els)) { if (e.getAttribute("Id") === embedId) return e as Element; }
        return null;
      })() : null;
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

      const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000;
      const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000;

      const element: ImageElement = {
        type: "image",
        relId: embedId,
        src: dataUri,
        position: { x, y },
        size: { width: cx, height: cy }
      };

      // Apply group coordinate transform if image is inside a group
      const groupXfrm = this.getGroupTransform(pic);
      if (groupXfrm) {
        const { gx, gy, cx0, cy0, scaleX, scaleY } = groupXfrm;
        const childRelX = (element.position.x - cx0) * scaleX;
        const childRelY = (element.position.y - cy0) * scaleY;
        element.position = { x: gx + childRelX, y: gy + childRelY };
        element.size = { width: element.size.width * scaleX, height: element.size.height * scaleY };
      }

      elements.push(element);
    }

    return elements;
  }

  /** Get group coordinate transform if the element is inside a <grpSp> */
  private static getGroupTransform(el: Element): {
    gx: number; gy: number; scaleX: number; scaleY: number;
    cx0: number; cy0: number;
  } | null {
    let parent = el.parentElement;
    while (parent) {
      if (parent.localName === "grpSp") {
        const grpSpPr = Array.from(parent.children).find(
          (c) => c.localName === "grpSpPr",
        ) as Element | undefined;
        if (!grpSpPr) return null;

        const grpXfrm = grpSpPr.getElementsByTagNameNS("*", "xfrm")[0];
        if (!grpXfrm) return null;

        const grpOff = grpXfrm.getElementsByTagNameNS("*", "off")[0];
        const grpExt = grpXfrm.getElementsByTagNameNS("*", "ext")[0];
        const chOff = grpXfrm.getElementsByTagNameNS("*", "chOff")[0];
        const chExt = grpXfrm.getElementsByTagNameNS("*", "chExt")[0];
        if (!grpOff || !grpExt || !chOff || !chExt) return null;

        const gx = XmlHelper.getAttrAsNumber(grpOff, "x");
        const gy = XmlHelper.getAttrAsNumber(grpOff, "y");
        const gw = XmlHelper.getAttrAsNumber(grpExt, "cx");
        const gh = XmlHelper.getAttrAsNumber(grpExt, "cy");
        const cx0 = XmlHelper.getAttrAsNumber(chOff, "x");
        const cy0 = XmlHelper.getAttrAsNumber(chOff, "y");
        const cw = XmlHelper.getAttrAsNumber(chExt, "cx");
        const ch = XmlHelper.getAttrAsNumber(chExt, "cy");

        return {
          gx, gy, cx0, cy0,
          scaleX: cw > 0 ? gw / cw : 1,
          scaleY: ch > 0 ? gh / ch : 1,
        };
      }
      if (parent.localName === "spTree") return null;
      parent = parent.parentElement;
    }
    return null;
  }

  /**
   * Normalizes a relative path from a slide rels file.
   * @param target Path from the relationship XML (e.g. "../media/image1.png")
   * @param basePath Base folder (e.g. "ppt/slides")
   * @returns Normalized path inside the zip (e.g. "ppt/media/image1.png")
   */
  private static normalizePath(target: string, basePath: string): string {
    const parts = (basePath + "/" + target).split("/");
    const resolved: string[] = [];
    for (const part of parts) {
      if (part === "..") resolved.pop();
      else if (part !== ".") resolved.push(part);
    }
    return resolved.join("/");
  }
}
