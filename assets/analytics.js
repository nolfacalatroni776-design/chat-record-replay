(function initAnalytics(root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.MessageRecordAnalytics = factory();
  }
})(typeof globalThis !== "undefined" ? globalThis : window, function buildAnalytics() {
  const REQUIRED_COLUMNS = ["消息内容", "发送用户", "用户角色", "群名", "项目ID", "发送时间"];
  const URL_RE = /https?:\/\/[^\s"'<>\\\]\}]+/gi;
  const IMAGE_EXTENSIONS = /\.(png|jpe?g|gif|webp|bmp|heic|svg)(\?|#|$)/i;
  const IMAGE_KEYWORD_PATTERNS = [
    /图片|照片|截图|附件|二维码|凭证|图像|文生图|生图|图[一二三四五六七八九十0-9]?/i,
    /image|photo|picture|screenshot|attachment|qr code|screen shot/i
  ];
  const SIGNAL_RULES = [
    {
      name: "问题/故障",
      pattern: /无法|不能|错误|失败|异常|问题|报错|卡住|没任务|没有任务|收不到|打不开|点不了|登录不了|\b(no jobs?|cannot|can't|error|failed|bug|issue|problem|not working|stuck|blocked)\b/i
    },
    {
      name: "求助/答疑",
      pattern: /请问|帮忙|怎么|如何|为什么|有人知道|麻烦|help|how to|anyone know|can you/i
    },
    {
      name: "任务/流程",
      pattern: /任务|项目|标注|转写|审核|提交|领取|进度|截止|培训|测试|task|project|annotation|transcription|submit|review|deadline|training/i
    },
    {
      name: "报酬/结算",
      pattern: /付款|结算|报酬|工资|收入|金额|钱包|银行卡|提现|payment|pay|paid|salary|bonus|wallet|bank/i
    },
    {
      name: "入群/二维码",
      pattern: /二维码|扫码|入群|加群|邀请|join group|qr code|scan|invite|lark/i
    },
    {
      name: "账号/权限",
      pattern: /账号|账户|登录|注册|权限|密码|邮箱|手机号|account|login|sign in|permission|password|email/i
    },
    {
      name: "公告/提醒",
      pattern: /通知|公告|提醒|注意|所有人|@所有人|warning|notice|announcement|reminder|everyone/i
    },
    {
      name: "确认/反馈",
      pattern: /收到|好的|ok|done|完成|确认|可以|thanks|thank you|confirmed|finished/i
    }
  ];

  function normalize(value) {
    return String(value ?? "").trim();
  }

  function uniq(values) {
    return Array.from(new Set(values.filter(Boolean)));
  }

  function countBy(items, keyFn) {
    const map = new Map();
    items.forEach((item) => {
      const key = keyFn(item) || "未识别";
      map.set(key, (map.get(key) || 0) + 1);
    });
    return Array.from(map, ([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
  }

  function sortTimeBuckets(rows) {
    return rows.sort((a, b) => {
      if (a.name === "未解析") return 1;
      if (b.name === "未解析") return -1;
      return a.name.localeCompare(b.name);
    });
  }

  function percent(value, total, digits = 1) {
    if (!total) return "0%";
    return `${((value / total) * 100).toFixed(digits)}%`;
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("zh-CN").format(Number(value || 0));
  }

  function isCellAddress(key) {
    return /^[A-Z]+[0-9]+$/.test(key);
  }

  function repairWorksheetRange(XLSX, worksheet) {
    let minR = Infinity;
    let minC = Infinity;
    let maxR = -1;
    let maxC = -1;

    Object.keys(worksheet).forEach((address) => {
      if (!isCellAddress(address)) return;
      const cell = XLSX.utils.decode_cell(address);
      minR = Math.min(minR, cell.r);
      minC = Math.min(minC, cell.c);
      maxR = Math.max(maxR, cell.r);
      maxC = Math.max(maxC, cell.c);
    });

    if (maxR >= 0) {
      worksheet["!ref"] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
    }
    return worksheet["!ref"] || "";
  }

  function headerStatus(rows) {
    const header = (rows[0] || []).slice(0, REQUIRED_COLUMNS.length).map(normalize);
    const missing = REQUIRED_COLUMNS.filter((column, index) => header[index] !== column);
    return { header, ok: missing.length === 0, missing };
  }

  function rowsFromWorksheet(XLSX, worksheet) {
    repairWorksheetRange(XLSX, worksheet);
    return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false, blankrows: false });
  }

  function recordsFromRows(rows) {
    const status = headerStatus(rows);
    if (!status.ok) {
      return { records: [], header: status.header, headerOk: false, headerMissing: status.missing };
    }

    const records = rows.slice(1)
      .map((row, index) => {
        const record = { __rowNumber: index + 2 };
        REQUIRED_COLUMNS.forEach((column, columnIndex) => {
          record[column] = normalize(row[columnIndex]);
        });
        return record;
      })
      .filter((record) => REQUIRED_COLUMNS.some((column) => record[column]));

    return { records, header: status.header, headerOk: true, headerMissing: [] };
  }

  function parseExcelDateSerial(value) {
    const serial = Number(value);
    if (!Number.isFinite(serial)) return null;
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    const fractionalDay = serial - Math.floor(serial) + 0.0000001;
    let totalSeconds = Math.floor(86400 * fractionalDay);
    const seconds = totalSeconds % 60;
    totalSeconds -= seconds;
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor(totalSeconds / 60) % 60;
    dateInfo.setUTCHours(hours, minutes, seconds);
    return dateInfo;
  }

  function parseDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    const text = normalize(value);
    if (!text) return null;
    if (/^\d+(\.\d+)?$/.test(text)) return parseExcelDateSerial(text);
    const normalized = text.replace(/\//g, "-").replace(" ", "T");
    const parsed = new Date(normalized);
    if (!Number.isNaN(parsed.getTime())) return parsed;
    const match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
    if (!match) return null;
    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4]),
      Number(match[5]),
      Number(match[6] || 0)
    );
  }

  function dateKey(date) {
    if (!date) return "未解析";
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function hourKey(date) {
    if (!date) return "未解析";
    return String(date.getHours()).padStart(2, "0");
  }

  function parseMessageParts(content) {
    const text = normalize(content);
    if (!text || !/^\s*[\[{]/.test(text)) return [];
    try {
      const parsed = JSON.parse(text);
      const list = Array.isArray(parsed) ? parsed : [parsed];
      return list
        .filter((item) => item && typeof item === "object")
        .map((item) => ({
          type: normalize(item.type).toUpperCase(),
          content: normalize(item.content)
        }));
    } catch {
      return [];
    }
  }

  function extractUrls(content) {
    return normalize(content).match(URL_RE) || [];
  }

  function isImageUrl(url) {
    const text = normalize(url);
    if (!text) return false;
    return IMAGE_EXTENSIONS.test(text) || /\/image(\?|#|$|\.)/i.test(text) || /path=.*\.(png|jpe?g|gif|webp|bmp|heic)/i.test(text);
  }

  function hasImageKeyword(content) {
    const text = normalize(content);
    return IMAGE_KEYWORD_PATTERNS.some((pattern) => pattern.test(text));
  }

  function imageUrlsFromMessage(content, parts) {
    const jsonImageUrls = parts
      .filter((part) => part.type === "IMAGE")
      .map((part) => part.content)
      .filter(Boolean);
    const linkedImageUrls = extractUrls(content).filter(isImageUrl);
    return uniq([...jsonImageUrls, ...linkedImageUrls]);
  }

  function textFromMessage(content, parts) {
    if (!parts.length) return normalize(content);
    const textParts = parts
      .filter((part) => part.type === "TEXT" || !part.type)
      .map((part) => part.content);
    return textParts.length ? textParts.join(" ").trim() : normalize(content);
  }

  function contentFormat(parts, text, urls) {
    if (!normalize(text) && parts.length) return "非文本结构";
    if (parts.length) return "结构化消息";
    if (urls.length) return "含链接文本";
    if (!normalize(text)) return "空内容";
    return "普通文本";
  }

  function detectSignals(record, text) {
    const haystack = `${record["消息内容"]} ${text} ${record["群名"]}`.toLowerCase();
    const signals = SIGNAL_RULES
      .filter((rule) => rule.pattern.test(haystack))
      .map((rule) => rule.name);
    return signals.length ? uniq(signals) : ["日常沟通"];
  }

  function classifyImageContent(record, text, imageAttachmentCount, imageKeywordOnly) {
    const lower = `${record["消息内容"]} ${text} ${record["群名"]}`.toLowerCase();
    if (/lark|qr code|scan|二维码|扫码|入群|加群|join the group|download lark/.test(lower)) return "入群/二维码";
    if (/http error|error|issue|problem|no jobs|bug|cannot|can't|点不了|不能|无法|收不了|截图|screenshot|screen shot/.test(lower)) return "问题反馈截图";
    if (/\$|payment|pay|bank|hour|hours|salary|wallet|结算|付款|报酬|收入|金额|凭证/.test(lower)) return "报酬/结算凭证";
    if (/warning|notice|everyone|@所有人|announcement|通知|谨防|受骗/.test(lower) && normalize(record["用户角色"]).includes("PM")) return "PM 通知素材";
    if (imageAttachmentCount > 0 && !text) return "纯图片附件";
    if (imageKeywordOnly) return "图片/附件线索";
    return "图片沟通";
  }

  function enrichRecord(record) {
    const parts = parseMessageParts(record["消息内容"]);
    const text = textFromMessage(record["消息内容"], parts);
    const urls = extractUrls(record["消息内容"]);
    const parsedDate = parseDate(record["发送时间"]);
    const signalTags = detectSignals(record, text);
    const imageUrls = imageUrlsFromMessage(record["消息内容"], parts);
    const imageAttachmentCount = imageUrls.length;
    const imageKeywordOnly = imageAttachmentCount === 0 && hasImageKeyword(`${record["消息内容"]} ${text}`);
    const isImageRelated = imageAttachmentCount > 0 || imageKeywordOnly;

    return {
      ...record,
      text,
      urls,
      imageUrls,
      imageAttachmentCount,
      hasImageAttachment: imageAttachmentCount > 0,
      imageKeywordOnly,
      isImageRelated,
      imageCategory: isImageRelated ? classifyImageContent(record, text, imageAttachmentCount, imageKeywordOnly) : "非图片附件",
      partTypes: uniq(parts.map((part) => part.type || "UNKNOWN")),
      contentFormat: contentFormat(parts, text, urls),
      signalTags,
      primarySignal: signalTags[0] || "日常沟通",
      parsedDate,
      day: dateKey(parsedDate),
      hour: hourKey(parsedDate)
    };
  }

  function topSignalFromMap(map) {
    const sorted = Array.from(map, ([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
    return sorted[0] || { name: "-", value: 0 };
  }

  function makeAggregate(records, key) {
    const map = new Map();
    records.forEach((record) => {
      const name = normalize(record[key]) || "未识别";
      if (!map.has(name)) {
        map.set(name, {
          name,
          messages: 0,
          pmMessages: 0,
          workerMessages: 0,
          users: new Set(),
          groups: new Set(),
          projects: new Set(),
          signals: new Map(),
          firstDate: null,
          lastDate: null
        });
      }
      const item = map.get(name);
      item.messages += 1;
      item.users.add(record["发送用户"]);
      item.groups.add(record["群名"]);
      item.projects.add(record["项目ID"]);
      if (record["用户角色"] === "IM_PM") item.pmMessages += 1;
      if (record["用户角色"] === "IM_WORKER") item.workerMessages += 1;
      item.signals.set(record.primarySignal, (item.signals.get(record.primarySignal) || 0) + 1);
      if (record.parsedDate) {
        if (!item.firstDate || record.parsedDate < item.firstDate) item.firstDate = record.parsedDate;
        if (!item.lastDate || record.parsedDate > item.lastDate) item.lastDate = record.parsedDate;
      }
    });

    return Array.from(map.values()).map((item) => {
      const topSignal = topSignalFromMap(item.signals);
      return {
        ...item,
        users: item.users.size,
        groups: item.groups.size,
        projects: item.projects.size,
        topSignal: topSignal.name,
        topSignalCount: topSignal.value,
        firstDay: dateKey(item.firstDate),
        lastDay: dateKey(item.lastDate)
      };
    }).sort((a, b) => b.messages - a.messages || a.name.localeCompare(b.name));
  }

  function makeSignalStats(records) {
    const map = new Map();
    records.forEach((record) => {
      const signal = record.primarySignal;
      if (!map.has(signal)) {
        map.set(signal, {
          name: signal,
          value: 0,
          pmMessages: 0,
          workerMessages: 0,
          users: new Set(),
          groups: new Set(),
          projects: new Set()
        });
      }
      const item = map.get(signal);
      item.value += 1;
      item.users.add(record["发送用户"]);
      item.groups.add(record["群名"]);
      item.projects.add(record["项目ID"]);
      if (record["用户角色"] === "IM_PM") item.pmMessages += 1;
      if (record["用户角色"] === "IM_WORKER") item.workerMessages += 1;
    });

    return Array.from(map.values()).map((item) => ({
      ...item,
      users: item.users.size,
      groups: item.groups.size,
      projects: item.projects.size
    })).sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
  }

  function makeImageAggregate(records, key) {
    const map = new Map();
    records.filter((record) => record.isImageRelated).forEach((record) => {
      const name = normalize(record[key]) || "未识别";
      if (!map.has(name)) {
        map.set(name, {
          name,
          imageRelatedMessages: 0,
          imageAttachmentMessages: 0,
          imageAttachments: 0,
          keywordOnlyMessages: 0,
          users: new Set(),
          groups: new Set(),
          projects: new Set()
        });
      }
      const item = map.get(name);
      item.imageRelatedMessages += 1;
      if (record.hasImageAttachment) item.imageAttachmentMessages += 1;
      if (record.imageKeywordOnly) item.keywordOnlyMessages += 1;
      item.imageAttachments += record.imageAttachmentCount;
      item.users.add(record["发送用户"]);
      item.groups.add(record["群名"]);
      item.projects.add(record["项目ID"]);
    });

    return Array.from(map.values()).map((item) => ({
      ...item,
      users: item.users.size,
      groups: item.groups.size,
      projects: item.projects.size
    })).sort((a, b) => b.imageRelatedMessages - a.imageRelatedMessages || b.imageAttachments - a.imageAttachments || a.name.localeCompare(b.name));
  }

  function analyzeRecords(records, source = {}) {
    const enriched = records.map(enrichRecord);
    const imageRelatedRecords = enriched.filter((record) => record.isImageRelated);
    const imageAttachmentRecords = enriched.filter((record) => record.hasImageAttachment);
    const dates = enriched.map((record) => record.parsedDate).filter(Boolean).sort((a, b) => a - b);
    const missingRequired = enriched.filter((record) => REQUIRED_COLUMNS.some((column) => !record[column])).length;
    const pmMessages = enriched.filter((record) => record["用户角色"] === "IM_PM").length;
    const workerMessages = enriched.filter((record) => record["用户角色"] === "IM_WORKER").length;
    const dayStats = sortTimeBuckets(countBy(enriched, (record) => record.day));
    const hourStats = sortTimeBuckets(countBy(enriched, (record) => record.hour));

    return {
      source,
      records: enriched,
      totals: {
        messages: enriched.length,
        groups: uniq(enriched.map((record) => record["群名"])).length,
        projects: uniq(enriched.map((record) => record["项目ID"])).length,
        users: uniq(enriched.map((record) => record["发送用户"])).length,
        pmMessages,
        workerMessages,
        imageRelatedMessages: imageRelatedRecords.length,
        imageAttachmentMessages: imageAttachmentRecords.length,
        imageAttachments: imageAttachmentRecords.reduce((sum, record) => sum + record.imageAttachmentCount, 0),
        imageKeywordOnlyMessages: imageRelatedRecords.filter((record) => record.imageKeywordOnly).length,
        imageGroups: uniq(imageRelatedRecords.map((record) => record["群名"])).length,
        imageProjects: uniq(imageRelatedRecords.map((record) => record["项目ID"])).length,
        imageUsers: uniq(imageRelatedRecords.map((record) => record["发送用户"])).length,
        validDates: dates.length,
        invalidDates: enriched.length - dates.length,
        missingRequired,
        dateStart: dates[0] || null,
        dateEnd: dates[dates.length - 1] || null,
        days: dayStats.filter((item) => item.name !== "未解析").length
      },
      roleStats: makeAggregate(enriched, "用户角色"),
      groupStats: makeAggregate(enriched, "群名"),
      projectStats: makeAggregate(enriched, "项目ID"),
      userStats: makeAggregate(enriched, "发送用户"),
      signalStats: makeSignalStats(enriched),
      imageCategoryStats: countBy(imageRelatedRecords, (record) => record.imageCategory),
      imageGroupStats: makeImageAggregate(enriched, "群名"),
      imageProjectStats: makeImageAggregate(enriched, "项目ID"),
      contentFormatStats: countBy(enriched, (record) => record.contentFormat),
      dayStats,
      hourStats,
      helpers: { percent, formatNumber }
    };
  }

  function buildExecutiveSummary(analysis) {
    const { totals, groupStats, projectStats, signalStats, dayStats } = analysis;
    if (!totals.messages) return "暂无数据。";
    const topGroup = groupStats[0];
    const topProject = projectStats[0];
    const topSignal = signalStats.find((item) => item.name !== "日常沟通") || signalStats[0];
    const peakDay = dayStats.filter((item) => item.name !== "未解析").slice().sort((a, b) => b.value - a.value)[0];
    const avgDaily = totals.days ? (totals.messages / totals.days).toFixed(0) : "0";
    const roleText = `PM ${formatNumber(totals.pmMessages)} 条，占 ${percent(totals.pmMessages, totals.messages)}；Worker ${formatNumber(totals.workerMessages)} 条，占 ${percent(totals.workerMessages, totals.messages)}。`;
    const groupText = topGroup ? `消息最集中在「${topGroup.name}」，${formatNumber(topGroup.messages)} 条。` : "";
    const projectText = topProject ? `Top 项目 ${topProject.name} 为 ${formatNumber(topProject.messages)} 条。` : "";
    const signalText = topSignal ? `主要内容信号为「${topSignal.name}」，命中 ${formatNumber(topSignal.value)} 条。` : "";
    const imageText = `附件图片线索 ${formatNumber(totals.imageRelatedMessages)} 条，占 ${percent(totals.imageRelatedMessages, totals.messages)}；其中实际图片附件消息 ${formatNumber(totals.imageAttachmentMessages)} 条、图片附件 ${formatNumber(totals.imageAttachments)} 个。`;
    const peakText = peakDay ? `日均 ${formatNumber(avgDaily)} 条，峰值 ${peakDay.name} 为 ${formatNumber(peakDay.value)} 条。` : "";
    return `本期共 ${formatNumber(totals.messages)} 条消息，覆盖 ${formatNumber(totals.groups)} 个群、${formatNumber(totals.projects)} 个项目、${formatNumber(totals.users)} 位用户。${roleText}${peakText}${groupText}${projectText}${signalText}${imageText}`;
  }

  function dateRangeText(analysis) {
    const start = analysis.totals.dateStart;
    const end = analysis.totals.dateEnd;
    if (!start || !end) return "-";
    return `${dateKey(start)} 至 ${dateKey(end)}`;
  }

  return {
    REQUIRED_COLUMNS,
    repairWorksheetRange,
    rowsFromWorksheet,
    recordsFromRows,
    analyzeRecords,
    buildExecutiveSummary,
    dateRangeText,
    percent,
    formatNumber,
    parseDate
  };
});
