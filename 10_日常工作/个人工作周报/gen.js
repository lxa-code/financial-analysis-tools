const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, ShadingType, WidthType, VerticalAlign, VerticalMergeType
} = require("docx");

// ============ 样式 ============
const FONT = "宋体";
const SZ = 24;
const HEADER_BG = "CED4D9";

const border = { style: BorderStyle.SINGLE, size: 2, color: "000000" };
const cb = { top: border, bottom: border, left: border, right: border };

function cP(text, bold) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: FONT, size: SZ, bold: !!bold })],
  });
}
function lP(text, bold) {
  return new Paragraph({
    children: [new TextRun({ text, font: FONT, size: SZ, bold: !!bold })],
  });
}

// 表头行（第一列合并col1+col2）
function sectionHeader(title) {
  return new TableRow({
    children: [
      new TableCell({ borders: cb, width: { size: 5592, type: WidthType.DXA }, columnSpan: 2, shading: { fill: HEADER_BG, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [cP(title, true)] }),
      new TableCell({ borders: cb, width: { size: 945, type: WidthType.DXA }, shading: { fill: HEADER_BG, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [cP("进度", true)] }),
      new TableCell({ borders: cb, width: { size: 994, type: WidthType.DXA }, shading: { fill: HEADER_BG, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [cP("工作成果", true)] }),
      new TableCell({ borders: cb, width: { size: 805, type: WidthType.DXA }, shading: { fill: HEADER_BG, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [cP("需协同单位", true)] }),
    ],
  });
}

// 分组行（合并5列显示任务名）
function groupRow(text) {
  return new TableRow({
    children: [
      new TableCell({ borders: cb, width: { size: 8336, type: WidthType.DXA }, columnSpan: 5, shading: { fill: "F2F2F2", type: ShadingType.CLEAR }, children: [lP(text, true)] }),
    ],
  });
}

// 数据行
function dataRow(content, progress, result, unit) {
  return new TableRow({
    children: [
      new TableCell({ borders: cb, width: { size: 1067, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, children: [cP(content)] }),
      new TableCell({ borders: cb, width: { size: 4525, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, children: [lP("")] }),
      new TableCell({ borders: cb, width: { size: 945, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, children: [cP(progress)] }),
      new TableCell({ borders: cb, width: { size: 994, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, children: [cP(result)] }),
      new TableCell({ borders: cb, width: { size: 805, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, children: [cP(unit)] }),
    ],
  });
}

// ============ 构建行 ============
const rows = [];

// 添加分组行和下面的数据行
function addGroup(name, items) {
  rows.push(groupRow(name));
  items.forEach(item => {
    rows.push(dataRow(item[0], item[1], item[2], item[3]));
  });
}

// ===== 周工作完成情况 =====
rows.push(sectionHeader("周工作完成情况"));

addGroup("任务一：物业商城系统及穿透式监管采购流程推进", [
  ["1.", "持续推进", "持续推进中", "综合管理部"],
  ["2.", "持续推进", "持续推进中", "综合管理部"],
]);

addGroup("任务二：税务相关事项沟通协调", [
  ["1.", "已完成", "完成发票额度调整沟通", "财务部"],
  ["2.", "已完成", "完成所得税汇算报告调整沟通", "财务部"],
]);

addGroup("任务三：运营管理体系梳理与数据支撑", [
  ["1.", "已完成", "初步梳理板块评价体系框架", "运营组"],
  ["2.", "已完成", "配合城市运营公司完成数据支撑", "运营组"],
  ["3.", "已完成", "配合完成月度重点工作梳理", "运营组"],
  ["4.", "进行中", "已启动组工作方案流程梳理", "运营组"],
]);

addGroup("问题分析与解决", [
  ["1.", "已解决", "多部门协调完成", "—"],
  ["2.", "跟进中", "持续推动审批流程", "—"],
]);

addGroup("体系", [
  ["1.", "持续优化", "完善运营管理体系框架", "运营组"],
]);

addGroup("要素", [
  ["1.", "持续提升", "数据支撑能力、AI知识库应用", "运营组"],
]);

// ===== 周工作计划 =====
rows.push(sectionHeader("周工作计划"));

addGroup("任务一：采购流程及系统建设", [
  ["1.", "持续推进", "推进物业商城系统采购", "综合管理部"],
  ["2.", "持续推进", "推进穿透式监管采购", "综合管理部"],
  ["3.", "计划开展", "听取光大银行、科蓝软件云收费系统汇报", "运营组"],
  ["4.", "计划开展", "梳理公司收发文流程", "综合管理部"],
]);

addGroup("任务二：运营管理体系梳理", [
  ["1.", "计划开展", "梳理成果体系，与AI知识库结合", "运营组"],
  ["2.", "持续推进", "深化板块评价机制建设", "运营组"],
]);

addGroup("任务三：税务及财务事项", [
  ["1.", "持续推进", "跟进发票额度调整后续", "财务部"],
]);

addGroup("问题分析与解决", [
  ["1.", "持续优化", "提升采购多部门协同效率", "—"],
]);

addGroup("体系", [
  ["1.", "计划开展", "运营管理体系框架标准化", "运营组"],
]);

addGroup("要素", [
  ["1.", "持续提升", "数据支撑、AI知识库应用能力", "运营组"],
]);

// ============ 构建文档 ============
const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
      },
    },
    children: [
      new Table({
        width: { size: 8336, type: WidthType.DXA },
        columnWidths: [1067, 4525, 945, 994, 805],
        rows,
      }),
    ],
  }],
});

const out = "e:\\BaiduSyncdisk\\CODE\\城发投资工作专区\\10_日常工作\\个人工作周报\\个人周报_20260525-0527.docx";
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(out, buf);
  console.log("Done:", out);
});
