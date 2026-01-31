export interface Position {
  /** Horizontal position in EMUs (English Metric Units) */
  x: number;

  /** Vertical position in EMUs */
  y: number;
}

export interface Size {
  /** Width in EMUs */
  width: number;

  /** Height in EMUs */
  height: number;
}

// ── Fill Types ──────────────────────────────────────────────────────────────

export interface SolidFill {
  type: "solid";
  color: string; // #RRGGBB
}

export interface GradientStop {
  offset: number; // 0–100 percentage
  color: string;  // #RRGGBB
}

export interface GradientFill {
  type: "gradient";
  gradientType: "linear" | "radial";
  stops: GradientStop[];
  /** Angle in CSS degrees (0 = bottom-to-top) for linear gradients */
  angle?: number;
}

export interface ImageFill {
  type: "image";
  src: string; // data URI or URL
}

export interface NoFill {
  type: "none";
}

/** Discriminated union for all fill types */
export type Fill = SolidFill | GradientFill | ImageFill | NoFill;

// ── Stroke Types ────────────────────────────────────────────────────────────

export interface ArrowHead {
  type?: string;
  w?: string;
  len?: string;
}

export interface StrokeStyle {
  color: string;  // #RRGGBB
  width: number;  // px
  dashStyle?: string; // OOXML preset dash name (dash, dot, dashDot, etc.)
  headEnd?: ArrowHead;
  tailEnd?: ArrowHead;
}

// ── Text Content (for text inside shapes) ───────────────────────────────────

export interface ShapeTextContent {
  html: string;
  font: { name: string; size: number; color: string };
  align?: { horizontal?: "left" | "center" | "right" | "justify"; vertical?: "top" | "middle" | "bottom" };
  padding?: { left: number; top: number; right: number; bottom: number };
}

// ── Element Types ───────────────────────────────────────────────────────────

export interface TextElement {
  type: "text";
  content: string;
  position: Position;
  size: Size;
  font: {
    name: string;
    size: number;
    color: string; // #RRGGBB
  };
  align?: {
    horizontal?: "left" | "center" | "right" | "justify";
    vertical?: "top" | "middle" | "bottom";
  };
  padding?: { left: number; top: number; right: number; bottom: number };
  html?: string;
}

export interface ImageElement {
  type: "image";
  relId: string;
  src: string;
  position: Position;
  size: Size;
}

export interface ShapeElement {
  type: "shape";
  shapeType: string;
  position: Position;
  size: Size;
  fill: Fill;
  stroke?: StrokeStyle;
  rotationDeg?: number;
  cornerRadiusPct?: number;
  textContent?: ShapeTextContent;
}

export interface BackgroundElement {
  type: "background";
  fill: Fill;
}

export interface TableCell {
  text: string;
  font?: { name?: string; size?: number; color?: string };
  align?: { horizontal?: "left" | "center" | "right" | "justify"; vertical?: "top" | "middle" | "bottom" };
  padding?: { left: number; top: number; right: number; bottom: number };
  fillColor?: string;
  colSpan?: number;
  rowSpan?: number;
  borders?: {
    top?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    right?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    bottom?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    left?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
  };
}

export interface TableRow {
  cells: TableCell[];
}

export interface TableElement {
  type: "table";
  position: Position;
  size: Size;
  columns: number[];
  rows: TableRow[];
  tableStyle?: { firstRow?: boolean; firstCol?: boolean; lastRow?: boolean; lastCol?: boolean; bandRow?: boolean; bandCol?: boolean };
  tableBorders?: {
    top?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    right?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    bottom?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    left?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    insideH?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    insideV?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
  };
  tableStyleId?: string;
  tableFillColor?: string;
  style?: {
    fills?: Partial<Record<
      "wholeTbl" | "band1H" | "band2H" | "band1V" | "band2V" | "firstRow" | "lastRow" | "firstCol" | "lastCol",
      string
    >>;
    fontColors?: Partial<Record<
      "wholeTbl" | "band1H" | "band2H" | "band1V" | "band2V" | "firstRow" | "lastRow" | "firstCol" | "lastCol",
      string
    >>;
  };
}

export type ChartType = "bar" | "column" | "line" | "pie" | "area" | "scatter";

export interface ChartSeries {
  name?: string;
  values?: number[];
  points?: { x: number; y: number }[];
  color?: string;
  valueFormat?: string;
}

export interface ChartElement {
  type: "chart";
  chartType: ChartType;
  position: Position;
  size: Size;
  categories: (string | number)[];
  series: ChartSeries[];
  palette?: string[];
  title?: string;
  showLegend?: boolean;
  showDataLabels?: boolean;
  stackedMode?: "none" | "stacked" | "percent";
  valueFormat?: string;
}

export type SlideElement =
  | TextElement
  | ImageElement
  | ShapeElement
  | BackgroundElement
  | TableElement
  | ChartElement;
