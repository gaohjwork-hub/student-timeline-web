const FEISHU_BASE = "https://open.feishu.cn/open-apis";

function json(statusCode, body) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    },
    body: JSON.stringify(body),
  };
}

function normalizeToken(input) {
  const text = String(input || "").trim();
  const match = text.match(/\/(?:sheets|spreadsheet)\/([A-Za-z0-9_-]+)/);
  if (match) return { type: "sheet", token: match[1] };
  const wikiMatch = text.match(/\/wiki\/([A-Za-z0-9_-]+)/);
  if (wikiMatch) {
    let sheetId = "";
    try {
      sheetId = new URL(text).searchParams.get("sheet") || "";
    } catch (_) {
      sheetId = "";
    }
    return { type: "wiki", token: wikiMatch[1], sheetId };
  }
  return { type: "sheet", token: text };
}

async function feishuFetch(path, options = {}) {
  const response = await fetch(`${FEISHU_BASE}${path}`, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.code !== 0) {
    const message = data.msg || data.error || `飞书接口 HTTP ${response.status}：${path}`;
    throw new Error(message);
  }
  return data;
}

async function getTenantAccessToken() {
  const appId = process.env.FEISHU_APP_ID;
  const appSecret = process.env.FEISHU_APP_SECRET;
  if (!appId || !appSecret) {
    throw new Error("Netlify 环境变量 FEISHU_APP_ID 或 FEISHU_APP_SECRET 未配置。");
  }
  const data = await feishuFetch("/auth/v3/tenant_access_token/internal", {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify({
      app_id: appId,
      app_secret: appSecret,
    }),
  });
  return data.tenant_access_token;
}

async function resolveSpreadsheetToken(resource, tenantAccessToken) {
  if (resource.type !== "wiki") return resource;
  const data = await feishuFetch(`/wiki/v2/spaces/get_node?token=${encodeURIComponent(resource.token)}`, {
    headers: { Authorization: `Bearer ${tenantAccessToken}` },
  });
  const node = data.data?.node || data.data;
  const objType = node?.obj_type || node?.objType;
  const objToken = node?.obj_token || node?.objToken;
  if (!objToken) {
    throw new Error("Wiki 节点没有返回对应的云文档 token，请确认链接指向电子表格。");
  }
  if (objType && !["sheet", "sheets", "spreadsheet"].includes(String(objType).toLowerCase())) {
    throw new Error(`这个 Wiki 节点不是电子表格，类型是 ${objType}。请打开真正的电子表格后复制 /sheets/ 链接。`);
  }
  return {
    type: "sheet",
    token: objToken,
    preferredSheetId: resource.sheetId,
  };
}

async function querySheets(token, tenantAccessToken) {
  const data = await feishuFetch(`/sheets/v3/spreadsheets/${token}/sheets/query`, {
    headers: { Authorization: `Bearer ${tenantAccessToken}` },
  });
  return (data.data?.sheets || []).map((sheet) => ({
    sheetId: sheet.sheet_id || sheet.sheetId,
    title: sheet.title,
    merges: sheet.merges || [],
  })).filter((sheet) => sheet.sheetId);
}

async function readValues(token, sheetId, range, tenantAccessToken) {
  const fullRange = `${sheetId}!${range || "A1:Z120"}`;
  const encodedRange = encodeURIComponent(fullRange);
  const data = await feishuFetch(`/sheets/v2/spreadsheets/${token}/values/${encodedRange}`, {
    headers: { Authorization: `Bearer ${tenantAccessToken}` },
  });
  return data.data?.valueRange?.values || [];
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return json(204, {});
  if (event.httpMethod !== "GET") return json(405, { error: "Only GET is supported." });

  try {
    const action = event.queryStringParameters?.action;
    const resource = normalizeToken(event.queryStringParameters?.token);
    const sheetId = event.queryStringParameters?.sheetId;
    const range = event.queryStringParameters?.range || "A1:Z120";
    if (!action) {
      return json(200, {
        ok: true,
        service: "feishu-sheets",
        hasAppId: Boolean(process.env.FEISHU_APP_ID),
        hasAppSecret: Boolean(process.env.FEISHU_APP_SECRET),
      });
    }
    if (!resource.token) return json(400, { error: "缺少飞书表格 token 或链接。" });

    const tenantAccessToken = await getTenantAccessToken();
    const spreadsheet = await resolveSpreadsheetToken(resource, tenantAccessToken);
    if (action === "sheets") {
      return json(200, {
        sheets: await querySheets(spreadsheet.token, tenantAccessToken),
        preferredSheetId: spreadsheet.preferredSheetId || "",
      });
    }
    if (action === "values") {
      if (!sheetId) return json(400, { error: "缺少 sheetId。" });
      const sheets = await querySheets(spreadsheet.token, tenantAccessToken);
      const sheet = sheets.find((item) => item.sheetId === sheetId);
      return json(200, {
        values: await readValues(spreadsheet.token, sheetId, range, tenantAccessToken),
        merges: sheet?.merges || [],
      });
    }
    return json(400, { error: "未知 action。" });
  } catch (error) {
    return json(500, { error: error.message || "飞书接口请求失败。" });
  }
};
