/* ============================================================
 * parser.js  —— Word(.docx) 解析逻辑 + 字段定义
 * ============================================================ */

/** 33 列字段定义（Excel 列顺序，A→AG）。全局共享给 app.js / exporter.js */
window.XF_FIELDS = [
  { col: 'A',  header: '信访形式',        field: 'fd_letter_type',       type: 'text' },
  { col: 'B',  header: '信访件编号',      field: 'fd_letter_no',         type: 'text' },
  { col: 'C',  header: '信访目的',        field: 'fd_letter_purpose',    type: 'text' },
  { col: 'D',  header: '信访日期',        field: 'fd_letter_date',       type: 'date' },
  { col: 'E',  header: '紧急程度',        field: 'fd_urgent_level',      type: 'text', optional: true },
  { col: 'F',  header: '登记机构',        field: 'fd_register_org',      type: 'text' },
  { col: 'G',  header: '联系人姓名',      field: 'fd_contact_name',      type: 'text' },
  { col: 'H',  header: '联系方式',        field: 'fd_contact_phone',     type: 'text' },
  { col: 'I',  header: '详细地址',        field: 'fd_detail_addr',       type: 'text' },
  { col: 'J',  header: '信访事项',        field: 'fd_letter_item',       type: 'textarea' },
  { col: 'K',  header: '主要诉求',        field: 'fd_main_demand',       type: 'textarea', optional: true },
  { col: 'L',  header: '备注',            field: 'fd_col_remark',        type: 'textarea', optional: true },
  { col: 'M',  header: '承办单位.名称*',  field: 'fd_deal_org.fdName',   type: 'text' },
  { col: 'N',  header: '抄送单位.名称*',  field: 'fd_copy_org.fdName',   type: 'text', optional: true },
  { col: 'O',  header: '证件号码',        field: 'fd_id_card',           type: 'text' },
  { col: 'P',  header: '信访人数',        field: 'fd_letter_person_num', type: 'text' },
  { col: 'Q',  header: '是否扬言',        field: 'fd_is_threat',         type: 'text' },
  { col: 'R',  header: '初件限办日期',    field: 'fd_deadline_date',     type: 'date' },
  { col: 'S',  header: '内容分类',        field: 'fd_content_type',      type: 'select',
    // 与 assets/template.xlsx 中 S 列的数据验证（下拉）值保持一致，另保留历史值“劳务纠纷”
    options: ['供水诉求','施工管理','管道不畅','管养问题','景区问题','井盖问题','拖欠农民工工资','拖欠工程款','其他','劳务纠纷'] },
  { col: 'T',  header: '初重',            field: 'fd_letter_level',      type: 'text' },
  { col: 'U',  header: '邮箱',            field: 'fd_email',             type: 'text', optional: true },
  { col: 'V',  header: '匿名信访',        field: 'fd_is_anonymous',      type: 'text' },
  { col: 'W',  header: '被反映人或单位',  field: 'fd_target_obj',        type: 'text', optional: true },
  { col: 'X',  header: '被反映人级别',    field: 'fd_target_level',      type: 'text' },
  { col: 'Y',  header: '问题发生地址',    field: 'fd_problem_addr',      type: 'text' },
  { col: 'Z',  header: '热点问题',        field: 'fd_hot_problem',       type: 'text', optional: true },
  { col: 'AA', header: '主办人员.名称*',  field: 'fd_col_47ha6p.fdName', type: 'text', blank: true },
  { col: 'AB', header: '关联案件.显示值', field: 'fd_relate_case.fdName',type: 'text', blank: true },
  { col: 'AC', header: '登记（受理）日期',field: 'fd_col_vt38ig',        type: 'text', blank: true },
  { col: 'AD', header: '是否受理',        field: 'fd_col_734136',        type: 'text', blank: true },
  { col: 'AE', header: '是否退件',        field: 'fd_col_ndvf9v',        type: 'text', blank: true },
  { col: 'AF', header: '退件理由',        field: 'fd_col_79j9xj',        type: 'text', blank: true },
  { col: 'AG', header: '文档标题',        field: 'fd_doc_subject',       type: 'text' }
];

/** Word 表格标签 → 内部临时键 */
window.XF_LABEL_MAP = {
  '信访目的': 'fd_letter_purpose',
  '信访日期': 'fd_letter_date',
  '登记机构': 'fd_register_org',
  '姓名': 'fd_contact_name',
  '证件号码': 'fd_id_card',
  '邮箱': 'fd_email',
  '住址': 'fd_addr',
  '手机': 'fd_contact_phone',
  '人数': 'fd_letter_person_num',
  '详细住址': 'fd_detail_addr_only',
  '被反映人或单位': 'fd_target_obj',
  '被反映人级别': 'fd_target_level',
  '问题发生地': 'fd_problem_addr',
  '热点问题': 'fd_hot_problem',
  '内容摘要': 'fd_content_summary',
  '投诉内容': 'fd_complaint_content', // 解析后丢弃
  '备注': 'fd_col_remark',
  '初件限办日期': 'fd_deadline_date',
  '初重': 'fd_letter_level',
  '承办单位': 'fd_deal_org_name',
  '抄送单位': 'fd_copy_org_name',
  '是否扬言': 'fd_is_threat',
  '匿名信访': 'fd_is_anonymous',
  '内容分类': 'fd_content_type',
  '紧急程度': 'fd_urgent_level'
};

window.XFParser = (function () {
  const clean = (s) => (s || '').replace(/[\u200b-\u200f\ufeff]/g, '').replace(/\s+/g, ' ').trim();
  const cleanMulti = (s) => (s || '').replace(/[\u200b-\u200f\ufeff]/g, '').replace(/[ \t]+/g, ' ').trim();

  /** 归一化日期为 yyyy-MM-dd；无法识别返回 null */
  function normalizeDate(raw) {
    const s = clean(raw);
    if (!s) return null;
    // 已是 yyyy-MM-dd
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return pad(m);
    // yyyy/MM/dd 或 yyyy.MM.dd
    m = s.match(/^(\d{4})[\/.](\d{1,2})[\/.](\d{1,2})$/);
    if (m) return pad(m);
    // yyyy年MM月dd日
    m = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
    if (m) return pad(m);
    // 提取任意日期片段
    m = s.match(/(\d{4})[-\/.年](\d{1,2})[-\/.月](\d{1,2})日?/);
    if (m) return pad(m);
    // Excel 序列号（5位数字）
    if (/^\d{5,6}$/.test(s)) {
      const d = new Date(Date.UTC(1899, 11, 30) + Number(s) * 86400000);
      if (!isNaN(d.getTime())) {
        return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()]
          .map((n) => String(n).padStart(2, '0')).join('-').replace(/^(\d{4})-/, '$1-');
      }
    }
    return null;
  }
  function pad(m) {
    return m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0');
  }

  /** 解析单个 .docx/.doc File，返回记录对象（含 _warnings / _file） */
  async function parseFile(file) {
    const raw = await file.arrayBuffer();
    const arrayBuffer = await ensureDocx(raw);
    const { value: html } = await mammoth.convertToHtml({ arrayBuffer });
    const doc = new DOMParser().parseFromString(html, 'text/html');
    return parseDocument(doc, file.name);
  }

  /**
   * 确保输入是可被 mammoth 解析的 .docx(zip) ArrayBuffer。
   * 省信访平台导出的“.doc”常为扁平化 OOXML(pkg:package XML)，需在浏览器内重组为 zip。
   */
  async function ensureDocx(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    // 已是 ZIP（真正 .docx）: 魔数 PK\x03\x04
    if (bytes[0] === 0x50 && bytes[1] === 0x4b) return arrayBuffer;
    // OLE2 二进制 .doc: 魔数 D0 CF 11 E0 —— mammoth 不支持
    if (bytes[0] === 0xd0 && bytes[1] === 0xcf) {
      throw new Error('旧版二进制 .doc 不支持，请用 Word 另存为 .docx 后重试');
    }
    let text;
    try { text = new TextDecoder('utf-8').decode(bytes); }
    catch (e) { throw new Error('无法识别的文件编码'); }
    if (text.indexOf('pkg:package') > -1) return flatOpcToDocx(text);
    // 纯 WordML(Word 2003 XML) 或其它 XML/HTML —— 不支持
    throw new Error('无法识别的文件格式（仅支持 .docx 或扁平化 OOXML 的 .doc）');
  }

  /** 将扁平化 OOXML(pkg:package) 重组为标准 .docx zip 的 ArrayBuffer */
  async function flatOpcToDocx(xmlText) {
    const dom = new DOMParser().parseFromString(xmlText, 'application/xml');
    if (dom.getElementsByTagName('parsererror').length) throw new Error('扁平 OOXML 解析失败');
    const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
    const parts = dom.getElementsByTagNameNS(PKG, 'part');
    if (!parts.length) throw new Error('未找到 pkg:part 节点');
    const zip = new JSZip();
    const serializer = new XMLSerializer();
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      let name = part.getAttributeNS(PKG, 'name') || part.getAttribute('pkg:name') || '';
      name = name.replace(/^\//, '');
      if (!name) continue;
      const binList = part.getElementsByTagNameNS(PKG, 'binaryData');
      const xmlList = part.getElementsByTagNameNS(PKG, 'xmlData');
      if (binList.length) {
        zip.file(name, binList[0].textContent.replace(/\s+/g, ''), { base64: true });
      } else if (xmlList.length) {
        const child = xmlList[0].firstElementChild;
        if (!child) continue;
        const s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + serializer.serializeToString(child);
        zip.file(name, s);
      }
    }
    return await zip.generateAsync({ type: 'arraybuffer' });
  }

  function parseDocument(doc, fileName) {
    const warnings = [];
    const tmp = {};

    // ---- 6.2 段落提取：信访形式 / 信访件编号 ----
    // 优先从段落(<p>/<h*>)逐个匹配，避免整篇 textContent 把表格正文一起吞入
    const parasOnly = Array.from(doc.querySelectorAll('p, h1, h2, h3, h4, li'))
      .filter((el) => !(el.closest && el.closest('table')))
      .map((el) => cleanMulti(el.textContent));
    const bodyText = parasOnly.join('\n');

    let formMatch = null, noMatch = null;
    for (const p of parasOnly) {
      if (!formMatch) formMatch = p.match(/信访形式[（(]([^）)]*)[）)]/);
      if (!noMatch) noMatch = p.match(/信访件编号[：:]\s*([A-Za-z0-9_\-]+|[^\s，,。；;：:]+?)(?=\s|$|[，,。；;])/);
      if (formMatch && noMatch) break;
    }
    // 回退：整篇文本（收紧编号字符集，避免贪婪吞并）
    if (!formMatch) formMatch = bodyText.match(/信访形式[（(]([^）)]*)[）)]/);
    if (!noMatch) noMatch = bodyText.match(/信访件编号[：:]\s*([A-Za-z0-9_\-]+)/);
    if (formMatch) tmp.fd_letter_type = clean(formMatch[1]);
    if (noMatch) tmp.fd_letter_no = clean(noMatch[1]);

    // ---- 6.3 表格提取 ----
    const tables = doc.querySelectorAll('table');
    tables.forEach((table) => {
      table.querySelectorAll('tr').forEach((tr) => {
        const cells = Array.from(tr.querySelectorAll('td, th'));
        for (let i = 0; i < cells.length - 1; i++) {
          const label = clean(cells[i].textContent).replace(/[：:]\s*$/, '');
          if (Object.prototype.hasOwnProperty.call(window.XF_LABEL_MAP, label)) {
            const key = window.XF_LABEL_MAP[label];
            const val = cleanMulti(cells[i + 1].textContent);
            if (tmp[key] === undefined || tmp[key] === '') tmp[key] = val;
          }
        }
      });
    });

    // ---- 6.4 内容摘要拆分 ----
    const summary = cleanMulti(tmp.fd_content_summary || '');
    if (summary) {
      const splitIndex = summary.indexOf('主要诉求：');
      if (splitIndex > -1) {
        let item = summary.substring(0, splitIndex).trim();
        const demand = summary.substring(splitIndex + '主要诉求：'.length).trim();
        item = item.replace(/^主要事实[：:]\s*/, '');
        tmp.fd_letter_item = item;
        tmp.fd_main_demand = demand;
      } else {
        tmp.fd_letter_item = summary.replace(/^主要事实[：:]\s*/, '');
        tmp.fd_main_demand = '';
        warnings.push('内容摘要中未找到"主要诉求："分隔符，主要诉求留空');
      }
    } else {
      tmp.fd_letter_item = '';
      tmp.fd_main_demand = '';
      warnings.push('内容摘要字段为空');
    }

    // ---- 组装最终记录 ----
    const record = {};
    record['fd_letter_type'] = tmp.fd_letter_type || '';
    record['fd_letter_no'] = tmp.fd_letter_no || '';
    record['fd_letter_purpose'] = tmp.fd_letter_purpose || '';
    record['fd_letter_date'] = normalizeDate(tmp.fd_letter_date) || (tmp.fd_letter_date || '');
    record['fd_urgent_level'] = tmp.fd_urgent_level || '';
    record['fd_register_org'] = tmp.fd_register_org || '';
    record['fd_contact_name'] = tmp.fd_contact_name || '';
    record['fd_contact_phone'] = tmp.fd_contact_phone || '';
    // 详细地址 = 住址 + ' ' + 详细住址
    const addr = [clean(tmp.fd_addr), clean(tmp.fd_detail_addr_only)].filter(Boolean).join(' ').trim();
    record['fd_detail_addr'] = addr;
    record['fd_letter_item'] = tmp.fd_letter_item || '';
    record['fd_main_demand'] = tmp.fd_main_demand || '';
    record['fd_col_remark'] = tmp.fd_col_remark || '';
    record['fd_deal_org.fdName'] = tmp.fd_deal_org_name || '';
    record['fd_copy_org.fdName'] = tmp.fd_copy_org_name || '';
    record['fd_id_card'] = tmp.fd_id_card || '';
    record['fd_letter_person_num'] = tmp.fd_letter_person_num || '';
    record['fd_is_threat'] = tmp.fd_is_threat || '';
    // 初件限办日期
    if (tmp.fd_deadline_date) {
      const nd = normalizeDate(tmp.fd_deadline_date);
      record['fd_deadline_date'] = nd || tmp.fd_deadline_date;
      if (!nd) warnings.push('初件限办日期格式异常：' + tmp.fd_deadline_date);
    } else {
      record['fd_deadline_date'] = '';
    }
    record['fd_content_type'] = tmp.fd_content_type || '';
    record['fd_letter_level'] = tmp.fd_letter_level || '';
    record['fd_email'] = tmp.fd_email || '';
    record['fd_is_anonymous'] = tmp.fd_is_anonymous || '';
    record['fd_target_obj'] = tmp.fd_target_obj || '';
    record['fd_target_level'] = tmp.fd_target_level || '';
    record['fd_problem_addr'] = tmp.fd_problem_addr || '';
    record['fd_hot_problem'] = tmp.fd_hot_problem || '';
    // 6 个留空字段
    record['fd_col_47ha6p.fdName'] = '';
    record['fd_relate_case.fdName'] = '';
    record['fd_col_vt38ig'] = '';
    record['fd_col_734136'] = '';
    record['fd_col_ndvf9v'] = '';
    record['fd_col_79j9xj'] = '';
    // 文档标题 = 信访件编号 + 姓名（不加空格）
    record['fd_doc_subject'] = (record['fd_letter_no'] || '') + (record['fd_contact_name'] || '');

    // ---- 日期格式提示 ----
    if (tmp.fd_letter_date && !/^\d{4}-\d{2}-\d{2}$/.test(record['fd_letter_date'])) {
      warnings.push('信访日期格式异常：' + tmp.fd_letter_date);
    }

    // ---- 内容分类校验 ----
    if (record['fd_content_type']) {
      const opts = window.XF_FIELDS.find((f) => f.field === 'fd_content_type').options;
      if (!opts.includes(record['fd_content_type'])) {
        warnings.push('内容分类字段不在可选范围，请手动修改：' + record['fd_content_type']);
      }
    } else {
      warnings.push('内容分类字段为空');
    }

    // ---- 6.6 缺失字段警告（非 optional 且非 blank 列）----
    window.XF_FIELDS.forEach((f) => {
      if (f.optional || f.blank) return;
      if (f.field === 'fd_letter_date' || f.field === 'fd_deadline_date' || f.field === 'fd_content_type') return; // 已单独处理
      if (!record[f.field]) {
        warnings.push(f.header + '字段为空');
      }
    });

    record._file = fileName;
    record._warnings = warnings;
    record._checked = true;
    return record;
  }

  return { parseFile, normalizeDate };
})();
