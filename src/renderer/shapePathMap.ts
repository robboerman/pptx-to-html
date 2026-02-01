/**
 * Maps known PowerPoint shape types to SVG path definitions.
 * Supports arrows, connectors, and geometric shapes.
 * @param type The shape type from the pptx.
 * @returns SVG path string to use in <path> or <polyline>, etc.
 */
export function getSvgPathForShape(type: string): string {
  switch (type) {
    // ▸ Basic Arrows
    case "rightArrow":
    case "arrow":
      return "POLYGON 0,25 70,25 70,0 100,50 70,100 70,75 0,75";

    case "leftArrow":
      return "POLYGON 100,25 30,25 30,0 0,50 30,100 30,75 100,75";

    case "leftRightArrow":
      return "POLYGON 0,50 30,0 30,25 70,25 70,0 100,50 70,100 70,75 30,75 30,100";

    case "triangle":
      return "POLYGON 50,0 100,100 0,100";

    case "star5":
      return "POLYGON 50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35";

    case "cloud":
      return "PATH M20,60 C10,60 10,40 25,40 C30,20 50,20 55,35 C70,30 80,40 80,50 C90,50 90,70 75,70 H25 Z";

    // ▸ Straight connector / line
    case "line":
    case "straightConnector1":
      return "LINE_ARROW 0,0 100,100";

    // ▸ Bent connectors
    case "bentConnector2":
      return "POLYLINE 0,50 50,50 50,100";

    case "bentConnector3":
      return "POLYLINE_ARROW 0,50 40,50 40,70 100,70";

    case "bentConnector4":
      return "POLYLINE 0,30 30,30 30,70 70,70 70,100";

    case "bentConnector5":
      return "POLYLINE 0,20 30,20 30,50 60,50 60,80 100,80";

    // ▸ Curved connectors
    case "curvedConnector2":
      return "PATH M0,50 Q50,0 100,50";

    case "curvedConnector3":
      return "PATH M0,50 Q25,0 50,50 Q75,100 100,50";

    case "curvedConnector4":
      return "PATH M0,40 Q20,0 40,40 Q60,80 80,40 Q90,20 100,40";

    case "curvedConnector5":
      return "PATH M0,50 Q20,20 40,50 Q60,80 80,50 Q90,40 100,50";

    // ▸ Notched, bent, and curved arrows
    case "bentArrow":
      return "POLYGON 0,0 70,0 70,30 100,30 50,100 50,30 0,30";

    case "notchedRightArrow":
      return "POLYGON 0,20 60,20 60,0 100,50 60,100 60,80 0,80";

    case "curvedRightArrow":
      return "PATH M0,50 Q50,0 100,50 Q50,100 0,50 Z";

    // ▸ Diagram / SmartArt Shapes
    case "homePlate":
      return "POLYGON 0,0 80,0 100,50 80,100 0,100";

    case "chevron":
      return "POLYGON 0,0 80,0 100,50 80,100 0,100 20,50";

    case "flowChartProcess":
      return "POLYGON 0,0 100,0 100,100 0,100";

    case "flowChartDecision":
      return "POLYGON 50,0 100,50 50,100 0,50";

    case "flowChartTerminator":
      return "PATH M20,0 H80 Q100,0 100,50 Q100,100 80,100 H20 Q0,100 0,50 Q0,0 20,0 Z";

    case "hexagon":
      return "POLYGON 25,0 75,0 100,50 75,100 25,100 0,50";

    case "parallelogram":
      return "POLYGON 20,0 100,0 80,100 0,100";

    case "trapezoid":
      return "POLYGON 20,0 80,0 100,100 0,100";

    case "pentagon":
      return "POLYGON 50,0 100,38 82,100 18,100 0,38";

    case "octagon":
      return "POLYGON 29,0 71,0 100,29 100,71 71,100 29,100 0,71 0,29";

    case "pie":
      return "PATH M50,50 L50,0 A50,50 0 1,1 49.99,0 Z";

    case "donut":
      return "PATH M50,0 A50,50 0 1,1 49.99,0 Z M50,25 A25,25 0 1,0 50.01,25 Z";

    case "blockArc":
      return "PATH M50,0 A50,50 0 1,1 49.99,0 L49.99,20 A30,30 0 1,0 50,20 Z";

    case "upArrow":
      return "POLYGON 25,100 25,40 0,40 50,0 100,40 75,40 75,100";

    case "downArrow":
      return "POLYGON 25,0 75,0 75,60 100,60 50,100 0,60 25,60";

    case "upDownArrow":
      return "POLYGON 50,0 100,30 70,30 70,70 100,70 50,100 0,70 30,70 30,30 0,30";

    case "stripedRightArrow":
      return "POLYGON 10,25 15,25 15,75 10,75 10,25 M20,25 70,25 70,0 100,50 70,100 70,75 20,75";

    case "doubleArrow":
      return "POLYGON 0,50 20,0 20,25 80,25 80,0 100,50 80,100 80,75 20,75 20,100";

    case "snip1Rect":
      return "POLYGON 0,0 85,0 100,15 100,100 0,100";

    case "snip2SameRect":
      return "POLYGON 15,0 85,0 100,15 100,100 0,100 0,15";

    case "round1Rect":
      return "PATH M15,0 H100 V100 H0 V15 Q0,0 15,0 Z";

    case "round2SameRect":
      return "PATH M15,0 H85 Q100,0 100,15 V100 H0 V15 Q0,0 15,0 Z";

    case "circularArrow":
      return "PATH M50,10 A40,40 0 1,1 15,60 L5,55 L20,75 L30,52 L20,48 A30,30 0 1,0 50,20 Z";

    case "gear6":
      return "PATH M43,2 L57,2 L60,15 L70,20 L82,12 L90,22 L80,33 L82,43 L95,48 L95,58 L82,62 L78,72 L88,82 L78,92 L68,82 L58,86 L57,98 L43,98 L40,86 L30,80 L18,88 L10,78 L20,68 L18,58 L5,53 L5,43 L18,38 L22,28 L12,18 L22,8 L32,18 L42,14 Z";

    case "gear9":
      return "PATH M47,2 L53,2 L55,12 L62,15 L70,8 L76,14 L70,22 L72,29 L82,30 L84,37 L75,42 L76,50 L86,53 L85,60 L75,60 L72,67 L79,75 L74,80 L66,74 L59,78 L58,88 L52,89 L48,80 L40,80 L35,89 L29,87 L32,78 L25,73 L17,79 L13,73 L20,66 L17,59 L7,58 L6,51 L16,48 L16,40 L6,36 L8,29 L18,30 L22,23 L16,15 L22,10 L30,17 L37,13 Z";

    // Default fallback (rectangular shape)
    default:
      return "POLYGON 0,0 100,0 100,100 0,100";
  }
}