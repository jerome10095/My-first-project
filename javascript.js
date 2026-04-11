javascript

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, LevelFormat
} = require('docx');
const fs = require('fs');

// ── helpers ──────────────────────────────────────────────────────────────────
const thin  = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
const allB  = { top: thin, bottom: thin, left: thin, right: thin };

const cell = (txt, w, isHeader = false, shade = false) =>
  new TableCell({
    borders: allB,
    width: { size: w, type: WidthType.DXA },
    shading: {
      fill: isHeader ? "1C3A5E" : (shade ? "EEF4FA" : "FFFFFF"),
      type: ShadingType.CLEAR
    },
    margins: { top: 90, bottom: 90, left: 130, right: 130 },
    children: [new Paragraph({
      children: [new TextRun({
        text: txt,
        bold: isHeader,
        size: isHeader ? 19 : 20,
        color: isHeader ? "FFFFFF" : "111111",
        font: "Calibri"
      })]
    })]
  });

const hRow  = (cells, widths) => new TableRow({ tableHeader: true,  children: cells.map((t,i) => cell(t, widths[i], true,  false)) });
const dRow  = (cells, widths, shade = false) => new TableRow({ children: cells.map((t,i) => cell(t, widths[i], false, shade)) });

const H1 = txt => new Paragraph({
  spacing: { before: 420, after: 160 },
  children: [new TextRun({ text: txt, bold: true, size: 30, color: "1C3A5E", font: "Calibri" })]
});
const H2 = txt => new Paragraph({
  spacing: { before: 280, after: 100 },
  children: [new TextRun({ text: txt, bold: true, size: 24, color: "2A6496", font: "Calibri" })]
});
const P = (txt, indent = false) => new Paragraph({
  spacing: { before: 80, after: 140 },
  indent: indent ? { left: 360 } : undefined,
  children: [new TextRun({ text: txt, size: 22, font: "Calibri" })]
});
const BP = (label, body) => new Paragraph({
  spacing: { before: 80, after: 140 },
  children: [
    new TextRun({ text: label + "  ", bold: true, size: 22, font: "Calibri" }),
    new TextRun({ text: body, size: 22, font: "Calibri" })
  ]
});
const BL = txt => new Paragraph({
  numbering: { reference: "bul", level: 0 },
  spacing: { before: 50, after: 50 },
  children: [new TextRun({ text: txt, size: 22, font: "Calibri" })]
});
const LINE = (color = "1C3A5E", sz = 6) => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: sz, color, space: 1 } },
  spacing: { before: 80, after: 80 },
  children: []
});
const GAP = () => new Paragraph({ spacing: { before: 50, after: 50 }, children: [] });

// ── document ─────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: "bul",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2022",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1260, bottom: 1080, left: 1260 }
      }
    },
    children: [

      // ── TITLE ──────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({
          text: "FINANCIAL RECOMMENDATIONS MEMORANDUM",
          bold: true, size: 34, color: "1C3A5E", font: "Calibri"
        })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({
          text: "N&D's Pizzeria — Pricing, Tariff Response & Operating Hours",
          size: 22, color: "555555", font: "Calibri"
        })]
      }),
      LINE("1C3A5E", 8),

      // memo header table
      new Table({
        width: { size: 9720, type: WidthType.DXA },
        columnWidths: [1700, 8020],
        rows: [
          dRow(["TO:", "Nicole & Danielle, Owners — N&D's Pizzeria"], [1700, 8020], false),
          dRow(["FROM:", "Financial Manager"], [1700, 8020], true),
          dRow(["DATE:", "April 10, 2026"], [1700, 8020], false),
          dRow(["SUBJECT:", "Pricing Strategy, Import Tariff Impacts & Extended Hours Analysis"], [1700, 8020], true),
        ]
      }),
      LINE("1C3A5E", 4), GAP(),

      // ── EXECUTIVE SUMMARY ──────────────────────────────────────
      H1("Executive Summary"),
      P("This memo answers three important questions for N&D's Pizzeria: (1) how to price meals given a possible import tariff on Italian ingredients, (2) whether to extend operating hours to midnight, and (3) how outside market conditions should shape everyday business decisions."),
      P("After reviewing the financial data, the income statement, and the competitor analysis from Success Marketing, the main recommendations are:"),
      BL("Keep the average price at $25 per meal. This gives the highest operating profit ($276,000 per year) and a strong profit margin of 31%, even though the $20 price level brings in slightly more total revenue."),
      BL("Move forward with extending hours, starting with a small pilot at a few locations. The estimated extra profit is about $70,000 per year once costs are covered."),
      BL("Do not immediately raise prices because of the tariff. The current margins are strong enough to absorb most of the cost increase. A small price adjustment of $0.50 to $1.00 per meal may be needed later if the tariff becomes permanent."),
      P("Overall, N&D's is in a healthy financial position. Making careful and well-informed decisions will help the business stay profitable and competitive."),
      GAP(), LINE("2A6496", 2),

      // ══════════════════════════════════════════════════════════
      H1("PART I — Market Forces and Trends"),

      // 1.1
      H2("1.1  Key Market Forces That Affect Pricing"),
      P("Market forces are conditions that influence how businesses set their prices. They are driven by the relationship between buyers and sellers in the market. For N&D's, four main forces matter the most."),

      BP("Supply and Demand:", "This is the most direct force. When more customers want to buy than the business can serve, prices can go up. When there are more products available than customers want, prices need to come down to attract buyers. For N&D's, the downtown office revitalization project is expected to bring around 1,000 new workers to the area, and universities are growing their student populations. Both of these trends increase the number of potential customers, which supports keeping prices steady or even raising them slightly in the future."),
      BP("Cost Structure:", "The price of any product must cover what it costs to make it. For N&D's, this includes ingredients, staff wages, rent for six locations, delivery vehicles, and equipment leases. The income statement shows that at $25 per meal, monthly operating costs are $52,000, which is the lowest of all price levels tested. This confirms that $25 is a financially sustainable price."),
      BP("Customer Perception of Value:"