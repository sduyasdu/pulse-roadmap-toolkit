/**
 * Generate a Yasdu-branded Pulse roadmap status report (.docx).
 *
 * Usage: node generate_report.js <input.json> <output.docx> <logo.png>
 *
 * See ../references/input_schema.md for the full shape of input.json.
 * All display strings (headings, labels, cell text) are expected to already
 * be in the target language — this script does not translate anything.
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, ImageRun, Header, Footer,
  PageNumber, VerticalAlign,
} = require("docx");

const [, , inputPath, outputPath, logoPath] = process.argv;
if (!inputPath || !outputPath || !logoPath) {
  console.error("Usage: node generate_report.js <input.json> <output.docx> <logo.png>");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(inputPath, "utf8"));
const logoBuffer = fs.readFileSync(logoPath);

const BLUE = "D85A28";
const DARK = "1E2433";
const GREY = "4B5468";
const LIGHT_BLUE_BG = "F7E8DA";
const ALT_ROW = "F7F8FA";
const BORDER_GREY = "D8DCE6";
const RED_TEXT = "991B1B";
const FONT = "Arial";

const PAGE_WIDTH_DXA = 12240; // US Letter
const MARGIN = 1440;
const USABLE_WIDTH = PAGE_WIDTH_DXA - 2 * MARGIN;

function cell(children, opts = {}) {
  return new TableCell({
    children,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shading ? { type: ShadingType.CLEAR, fill: opts.shading } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
  });
}

function headerRow(labels, widths) {
  return new TableRow({
    tableHeader: true,
    children: labels.map((label, i) =>
      cell(
        [new Paragraph({ children: [new TextRun({ text: label, bold: true, color: "FFFFFF", size: 19, font: FONT })] })],
        { width: widths[i], shading: BLUE }
      )
    ),
  });
}

function subtaskParagraphs(subtasks, colorOverride) {
  const color = colorOverride || DARK;
  if (!subtasks || subtasks.length === 0) {
    return [new Paragraph({ children: [new TextRun({ text: "—", size: 19, color, font: FONT })] })];
  }
  return subtasks.map(
    (label, i) =>
      new Paragraph({
        spacing: i > 0 ? { before: 40 } : undefined,
        children: [new TextRun({ text: `• ${label}`, size: 19, color, font: FONT })],
      })
  );
}

function dataRow(cells, widths, altShade, colorOverride) {
  return new TableRow({
    cantSplit: true,
    children: cells.map((c, i) => {
      // A cell value can be a plain string, or { subtasks: [...] } for a bulleted subtasks column
      let paragraphs;
      if (Array.isArray(c)) {
        paragraphs = c; // already an array of Paragraphs
      } else if (c && typeof c === "object" && c.subtasks !== undefined) {
        paragraphs = subtaskParagraphs(c.subtasks, colorOverride);
      } else {
        paragraphs = [new Paragraph({ children: [new TextRun({ text: String(c), size: 19, color: colorOverride || DARK, font: FONT })] })];
      }
      return cell(paragraphs, { width: widths[i], shading: altShade ? ALT_ROW : undefined });
    }),
  });
}

function buildTable(columns, rows) {
  const headers = columns.map((c) => c.label);
  const widths = columns.map((c) => Math.round(USABLE_WIDTH * c.width));
  return new Table({
    width: { size: USABLE_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      headerRow(headers, widths),
      ...rows.map((r, idx) => dataRow(r.cells, widths, idx % 2 === 1, r.color)),
    ],
  });
}

function sectionHeading(text, color) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 140 },
    children: [new TextRun({ text, bold: true, color: color || BLUE, size: 30, font: FONT })],
  });
}

function epicHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 100 },
    children: [new TextRun({ text, bold: true, color: "123359", size: 23, font: FONT })],
  });
}

function introPara(text) {
  return new Paragraph({
    spacing: { after: 180 },
    children: [new TextRun({ text, size: 20, color: GREY, font: FONT })],
  });
}

// ---------- Build section content from data.sections[] ----------

function buildSection(section) {
  const children = [];
  const headingColor = section.colorOverride || BLUE;
  const headingText = section.icon ? `${section.icon} ${section.heading}` : section.heading;
  children.push(sectionHeading(headingText, headingColor));
  if (section.introText) children.push(introPara(section.introText));

  for (const group of section.epics) {
    if (group.name) children.push(epicHeading(group.name));
    const rows = group.rows.map((r) => ({
      cells: r.cells,
      color: section.colorOverride || undefined,
    }));
    children.push(buildTable(section.columns, rows));
  }
  return children;
}

// ---------- Assemble document ----------

const bodyChildren = [
  new Paragraph({
    spacing: { before: 100, after: 40 },
    children: [new TextRun({ text: data.meta.title, bold: true, color: BLUE, size: 56, font: FONT })],
  }),
  new Paragraph({
    spacing: { after: 240 },
    children: [new TextRun({ text: data.meta.subtitle, size: 24, color: GREY, font: FONT })],
  }),
  introPara(data.meta.intro),
];

for (const section of data.sections) {
  bodyChildren.push(...buildSection(section));
}

if (data.keyRisks) {
  bodyChildren.push(
    new Table({
      width: { size: USABLE_WIDTH, type: WidthType.DXA },
      columnWidths: [USABLE_WIDTH],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: USABLE_WIDTH, type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: LIGHT_BLUE_BG },
              borders: {
                left: { style: BorderStyle.SINGLE, size: 24, color: "123359" },
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              },
              margins: { top: 160, bottom: 160, left: 220, right: 220 },
              children: [
                new Paragraph({
                  spacing: { after: 60 },
                  children: [new TextRun({ text: data.meta.keyRisksLabel || "KEY RISKS", bold: true, color: "123359", size: 18, font: FONT })],
                }),
                new Paragraph({
                  children: [new TextRun({ text: data.keyRisks, size: 20, color: "123359", font: FONT })],
                }),
              ],
            }),
          ],
        }),
      ],
    })
  );
  bodyChildren.push(new Paragraph({ spacing: { before: 120 }, children: [] }));
}

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONT } },
      heading1: { run: { font: FONT } },
      heading2: { run: { font: FONT } },
    },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: PAGE_WIDTH_DXA, height: 15840 },
          margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER_GREY, space: 6 } },
              children: [new ImageRun({ data: logoBuffer, transformation: { width: 80, height: 23 }, type: "png" })],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({ text: `${data.meta.footerText || data.meta.title}   `, size: 16, color: "8E96A8", font: FONT }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "8E96A8", font: FONT }),
              ],
            }),
          ],
        }),
      },
      children: bodyChildren,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outputPath, buf);
  console.log("Wrote", outputPath);
});
