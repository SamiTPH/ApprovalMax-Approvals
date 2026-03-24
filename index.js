require("dotenv").config();
const axios = require("axios");

// ==============================
// 🔐 REFRESH APPROVALMAX TOKEN
// ==============================
async function refreshApprovalMaxToken() {
  try {
    const params = new URLSearchParams();
    params.append("grant_type", "refresh_token");
    params.append("refresh_token", process.env.REFRESH_TOKEN);
    params.append("client_id", process.env.APPROVALMAX_CLIENT_ID);
    params.append("client_secret", process.env.APPROVALMAX_CLIENT_SECRET);

    const res = await axios.post(
      "https://identity.approvalmax.com/connect/token",
      params
    );

    const { access_token, refresh_token } = res.data;

    console.log("🔄 ApprovalMax token refreshed");

    // ✅ Keep in memory ONLY (Railway-safe)
    process.env.ACCESS_TOKEN = access_token;
    process.env.REFRESH_TOKEN = refresh_token;

    return access_token;

  } catch (err) {
    console.error("❌ Token refresh failed:", err.response?.data || err.message);
    throw err;
  }
}

// ==============================
// 📥 FETCH APPROVALS
// ==============================
async function fetchApprovals() {
  try {
    const res = await axios.get(
      `https://public-api.approvalmax.com/api/v1/companies/${process.env.COMPANY_ID}/standalone/documents`,
      {
        headers: {
          Authorization: `Bearer ${process.env.ACCESS_TOKEN}`,
          Accept: "application/json"
        }
      }
    );

    return res.data.payload || [];

  } catch (err) {
    if (err.response?.status === 401) {
      console.log("⚠️ Token expired → refreshing...");
      const newToken = await refreshApprovalMaxToken();

      const retry = await axios.get(
        `https://public-api.approvalmax.com/api/v1/companies/${process.env.COMPANY_ID}/standalone/documents`,
        {
          headers: {
            Authorization: `Bearer ${newToken}`,
            Accept: "application/json"
          }
        }
      );

      return retry.data.payload || [];
    }

    throw err;
  }
}

// ==============================
// 🧠 CLEAN DESCRIPTION
// ==============================
function cleanDescription(html) {
  if (!html) return "";

  return html
    .replace(/<\/p>/g, "\n")
    .replace(/<br\s*\/?>/g, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/\n+/g, "\n")
    .trim();
}

// ==============================
// 🧠 GET DEPARTMENT
// ==============================
function getDepartment(doc) {
  const fields = doc.additionalInformation || [];

  const dept = fields.find(
    f => f.additionalFieldName === "Department"
  );

  return dept?.value || null;
}

// ==============================
// 📊 FORMAT ROW
// ==============================
function formatRow(doc) {
  return [
    doc.requestId,
    doc.documentName || "",
    cleanDescription(doc.description),
    doc.friendlyName || "",
    doc.requestStatus || "",
    getDepartment(doc) || "",
    doc.createdAt || "",
    doc.modifiedAt || "",
    doc.decisionDate || ""
  ];
}

// ==============================
// 🔐 GRAPH TOKEN
// ==============================
async function getGraphToken() {
  const params = new URLSearchParams();
  params.append("client_id", process.env.MS_CLIENT_ID);
  params.append("client_secret", process.env.MS_CLIENT_SECRET);
  params.append("scope", "https://graph.microsoft.com/.default");
  params.append("grant_type", "client_credentials");

  const res = await axios.post(
    `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`,
    params
  );

  return res.data.access_token;
}

// ==============================
// 📥 GET EXCEL ROWS
// ==============================
async function getExcelRows(token) {
  const url =
    `https://graph.microsoft.com/v1.0/users/${process.env.EXCEL_USER}/drive/root:${process.env.EXCEL_FILE_PATH}:/workbook/tables/${process.env.EXCEL_TABLE_NAME}/rows`;

  const res = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  return res.data.value || [];
}

// ==============================
// 🔁 BUILD MAP
// ==============================
function buildMap(rows) {
  const map = {};

  rows.forEach(row => {
    const values = row.values[0];
    map[values[0]] = values;
  });

  return map;
}

// ==============================
// 🔄 PREPARE ROWS
// ==============================
function prepareRows(apiData, excelMap) {
  const rowsToInsert = [];

  apiData.forEach(doc => {
    const dept = getDepartment(doc);

    if (dept !== "Sales & Pre-Sales") return;

    const existing = excelMap[doc.requestId];
    const newRow = formatRow(doc);

    if (!existing) {
      rowsToInsert.push(newRow);
    } else if (existing[4] !== doc.requestStatus) {
      rowsToInsert.push(newRow);
    }
  });

  return rowsToInsert;
}

// ==============================
// ➕ ADD ROWS
// ==============================
async function addRows(token, rows) {
  if (rows.length === 0) {
    console.log("No updates needed");
    return;
  }

  const url =
    `https://graph.microsoft.com/v1.0/users/${process.env.EXCEL_USER}/drive/root:${process.env.EXCEL_FILE_PATH}:/workbook/tables/${process.env.EXCEL_TABLE_NAME}/rows/add`;

  await axios.post(url, { values: rows }, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    }
  });

  console.log(`✅ Added ${rows.length} rows`);
}

// ==============================
// 🚀 MAIN
// ==============================
async function main() {
  try {
    console.log("Fetching approvals...");
    const approvals = await fetchApprovals();

    console.log("Getting Graph token...");
    const token = await getGraphToken();

    console.log("Reading Excel...");
    const excelRows = await getExcelRows(token);

    const excelMap = buildMap(excelRows);

    console.log("Comparing...");
    const rows = prepareRows(approvals, excelMap);

    console.log("Rows to insert:", rows.length);

    await addRows(token, rows);

    console.log("✅ Sync complete");

  } catch (err) {
    console.error("❌ ERROR:", err.response?.data || err.message);
  }
}

main();