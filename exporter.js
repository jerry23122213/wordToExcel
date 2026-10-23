/* ============================================================
 * exporter.js —— 在「模板 xlsx 包」层面追加数据行
 *
 * 设计要点（为什么不用 XLSX.read → XLSX.write 往返转换）：
 *   原模板是 .xls（OLE2 二进制），XLSX 库读不出它的样式，
 *   往返写出的 .xlsx 会丢掉框线 / 列宽 / 合并单元格 / 数据验证，
 *   平台导入时报「格式错误」。
 *
 * 现在改为：
 *   1) 内置一份由原 .xls 另存得到的 template.xlsx（base64）；
 *   2) 用 JSZip 打开它，只在 xl/worksheets/sheet*.xml 的 sheetData
 *      里追加数据行（复用模板第 3 行的单元格样式，框线自动继承）；
 *   3) 其余全部部件原样写回 —— styles.xml / cols / mergeCells /
 *      dataValidations / 「样例」sheet 均与模板完全一致。
 * ============================================================ */
window.XFExporter = (function () {
  const TARGET_SHEET = '省一体化平台登记表';
  const SAMPLE_SHEET = '样例';
  const DATA_START_ROW = 3; // 数据从第 3 行（1 基）开始写，覆盖原占位行

  /* ---------------- 基础工具 ---------------- */

  // XML 文本转义 + 去掉 XML 非法控制字符
  function xmlEscape(value) {
    return String(value)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // 单元格文本归一化：平台导入只支持单行文本，换行/制表符压成空格
  function normalizeText(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/\r\n|\r|\n|\t/g, ' ').replace(/\s{2,}/g, ' ').trim();
  }

  function isNumeric(text) {
    return /^-?\d+(\.\d+)?$/.test(text);
  }

  function timestamp() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return '' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) +
      p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds());
  }

  /* ---------------- 读取模板 zip ---------------- */

  async function loadTemplateZip() {
    const JSZipLib = window.JSZip;
    if (!JSZipLib) throw new Error('缺少 JSZip 依赖，请检查网络是否能加载 CDN 资源');

    const b64 = (window.XF_TEMPLATE_BASE64 || '').replace(/[\s\r\n]/g, '');
    if (b64) return await JSZipLib.loadAsync(b64, { base64: true });

    // 兜底：直接取 assets/template.xlsx（需 http(s) 环境，file:// 下会失败）
    try {
      const resp = await fetch('assets/template.xlsx', { cache: 'no-store' });
      if (resp && resp.ok) return await JSZipLib.loadAsync(await resp.arrayBuffer());
    } catch (e) { /* ignore */ }
    throw new Error('无法读取内置导入模板，请确认 assets/template-data.js 已正确加载');
  }

  // 依据 workbook.xml + rels 定位目标 sheet 的 xml 路径
  async function resolveSheetPath(zip, sheetName) {
    const wbFile = zip.file('xl/workbook.xml');
    if (!wbFile) throw new Error('模板文件结构异常：缺少 xl/workbook.xml');
    const wbXml = await wbFile.async('string');

    const sheetTags = wbXml.match(/<sheet\b[^>]*>/g) || [];
    let rid = null;
    let sheetIndex = -1;
    sheetTags.forEach((tag, i) => {
      const name = (tag.match(/\bname="([^"]*)"/) || [])[1];
      if (name === sheetName) {
        rid = (tag.match(/\br:id="([^"]*)"/) || [])[1];
        sheetIndex = i;
      }
    });
    if (sheetIndex === -1) throw new Error('模板缺少工作表「' + sheetName + '」');

    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (relsFile && rid) {
      const relsXml = await relsFile.async('string');
      const rels = (relsXml.match(/<Relationship\b[^>]*>/g) || []).map((t) => ({
        id: (t.match(/\bId="([^"]*)"/) || [])[1],
        target: (t.match(/\bTarget="([^"]*)"/) || [])[1]
      }));
      const hit = rels.find((r) => r.id === rid);
      if (hit && hit.target) {
        const target = hit.target.replace(/^\/+/, '').replace(/^\.\//, '');
        const path = target.indexOf('xl/') === 0 ? target : 'xl/' + target;
        if (zip.file(path)) return path;
      }
    }

    const guess = 'xl/worksheets/sheet' + (sheetIndex + 1) + '.xml';
    if (zip.file(guess)) return guess;
    throw new Error('无法定位模板工作表文件');
  }

  /* ---------------- 解析 / 生成表格行 ---------------- */

  // 从模板的占位数据行里取出「每列对应样式索引」，数据行复用这些样式（含框线）
  function parseStyleMap(rowXml) {
    const map = {};
    const cellRe = /<c\b([^>]*?)\/?>/g;
    let m;
    while ((m = cellRe.exec(rowXml)) !== null) {
      const attrs = m[1];
      const ref = (attrs.match(/\br="([A-Z]+)\d+"/) || [])[1];
      const style = (attrs.match(/\bs="(\d+)"/) || [])[1];
      if (ref && style !== undefined) map[ref] = Number(style);
    }
    return map;
  }

  function buildRow(rowNum, record, styleMap, fields) {
    let cells = '';
    for (let i = 0; i < fields.length; i += 1) {
      const field = fields[i];
      const ref = field.col + rowNum;
      const sAttr = styleMap[field.col] !== undefined ? ' s="' + styleMap[field.col] + '"' : '';
      const text = normalizeText(record ? record[field.field] : '');

      if (text === '') {
        // 空值也写一个带样式的空单元格，保证整行框线连续
        if (sAttr) cells += '<c r="' + ref + '"' + sAttr + '/>';
        continue;
      }
      if (field.field === 'fd_letter_person_num' && isNumeric(text)) {
        cells += '<c r="' + ref + '"' + sAttr + '><v>' + text + '</v></c>';
        continue;
      }
      cells += '<c r="' + ref + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' +
        xmlEscape(text) + '</t></is></c>';
    }
    return '<row r="' + rowNum + '" spans="1:' + fields.length + '" x14ac:dyDescent="0.35">' + cells + '</row>';
  }

  // 只替换 sheetData 与 dimension，其余 XML 原样保留
  function patchSheetXml(xml, records) {
    const fields = Array.isArray(window.XF_FIELDS) ? window.XF_FIELDS : [];
    if (!fields.length) throw new Error('缺少字段定义（XF_FIELDS）');

    const sdMatch = xml.match(/<sheetData\s*\/>|<sheetData>[\s\S]*?<\/sheetData>/);
    if (!sdMatch) throw new Error('模板工作表结构异常：未找到 sheetData');
    const sheetData = sdMatch[0];

    const rowRe = /<row\b[^>]*\/>|<row\b[\s\S]*?<\/row>/g;
    const rows = sheetData.match(rowRe) || [];

    // 保留第 3 行之前的行（说明行、表头行），其余占位行移除
    const keepRows = [];
    let templateRowXml = null;
    rows.forEach((rowXml) => {
      const rNum = Number((rowXml.match(/\br="(\d+)"/) || [])[1] || 0);
      if (!rNum) { keepRows.push(rowXml); return; }
      if (rNum < DATA_START_ROW) keepRows.push(rowXml);
      else if (rNum === DATA_START_ROW) templateRowXml = rowXml;
    });
    // 没有占位数据行时不复用任何单元格样式（避免误用表头填充色）
    const styleMap = templateRowXml ? parseStyleMap(templateRowXml) : {};
    const dataRows = records.map((rec, i) => buildRow(DATA_START_ROW + i, rec, styleMap, fields));
    const lastRow = DATA_START_ROW + records.length - 1;

    let out = xml.replace(sheetData, '<sheetData>' + keepRows.join('') + dataRows.join('') + '</sheetData>');

    // 同步 dimension（仅提示用，但不更新会导致部分软件按旧范围渲染）
    out = out.replace(/<dimension\b[^>]*\/>/, (tag) => {
      const ref = (tag.match(/\bref="([^"]*)"/) || [])[1];
      const start = ref ? ref.split(':')[0] : 'A1';
      const endCol = fields[fields.length - 1].col;
      return tag.replace(/\bref="[^"]*"/, 'ref="' + start + ':' + endCol + lastRow + '"');
    });

    return out;
  }

  /* ---------------- 导出入口 ---------------- */

  async function exportRecords(records) {
    const list = Array.isArray(records) ? records : [];
    if (!list.length) throw new Error('没有可导出的记录');

    const zip = await loadTemplateZip();
    const sheetPath = await resolveSheetPath(zip, TARGET_SHEET);
    if (!zip.file('xl/styles.xml')) throw new Error('模板文件结构异常：缺少 styles.xml');

    const original = await zip.file(sheetPath).async('string');
    zip.file(sheetPath, patchSheetXml(original, list));

    // 「样例」sheet 及其他部件一律不写，保持与模板逐字节一致
    // 清掉 JSZip 自动补出的目录条目（Excel 原生 zip 不含 "xl/" 这类条目）
    Object.keys(zip.files).forEach((key) => {
      const entry = zip.files[key];
      if (entry && entry.dir) delete zip.files[key];
    });

    const fileName = '信访导入_' + timestamp() + '.xlsx';
    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      createFolders: false // 不写入目录条目，保持与 Excel 原生 zip 结构一致
    });

    if (typeof window.saveAs === 'function') {
      window.saveAs(blob, fileName);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    }

    return { fileName: fileName, count: list.length };
  }

  return {
    exportRecords,
    // 便于调试
    _patchSheetXml: patchSheetXml,
    TARGET_SHEET,
    SAMPLE_SHEET
  };
})();
