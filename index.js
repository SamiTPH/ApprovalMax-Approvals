require("dotenv").config();
const axios = require("axios");
const { Pool } = require("pg");

// ==============================
// DB CONNECTION
// ==============================
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

// ==============================
// TOKEN HANDLING
// ==============================
async function getStoredTokens() {
  const res = await pool.query("SELECT * FROM tokens LIMIT 1");
  return res.rows[0];
}

async function saveTokens(access_token, refresh_token) {
  await pool.query(`
    UPDATE tokens
    SET access_token=$1, refresh_token=$2, updated_at=NOW()
    WHERE id = (SELECT id FROM tokens LIMIT 1)
  `, [access_token, refresh_token]);
}

async function refreshToken() {
  const stored = await getStoredTokens();

  const params = new URLSearchParams();
  params.append("grant_type", "refresh_token");
  params.append("refresh_token", stored.refresh_token);
  params.append("client_id", process.env.APPROVALMAX_CLIENT_ID);
  params.append("client_secret", process.env.APPROVALMAX_CLIENT_SECRET);

  const res = await axios.post(
    "https://identity.approvalmax.com/connect/token",
    params
  );

  const { access_token, refresh_token } = res.data;

  console.log("🔄 Token refreshed");

  await saveTokens(access_token, refresh_token);

  return access_token;
}

// ==============================
// FETCH DOCUMENTS (CORRECT PAGINATION)
// ==============================
async function fetchAllDocuments(token) {
  let allDocs = [];
  let continuationToken = null;

  console.log("📄 Fetching documents...");

  do {
    const url = new URL(
      `https://public-api.approvalmax.com/api/v1/companies/${process.env.COMPANY_ID}/standalone/documents`
    );

    url.searchParams.append("limit", "100");

    if (continuationToken) {
      url.searchParams.append("continuationToken", continuationToken);
    }

    const res = await axios.get(url.toString(), {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json"
      }
    });

    const data = res.data;

    const batch = data.items || [];
    allDocs.push(...batch);

    console.log(`📄 Batch fetched: ${batch.length}`);

    continuationToken = data.continuationToken;

  } while (continuationToken);

  console.log(`✅ TOTAL DOCUMENTS: ${allDocs.length}`);

  return allDocs;
}

// ==============================
// FETCH USERS (FIXED)
// ==============================
async function fetchUsers(token) {
  console.log("👤 Fetching users...");

  const res = await axios.get(
    `https://public-api.approvalmax.com/api/v1/companies/${process.env.COMPANY_ID}/userProfiles`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json"
      }
    }
  );

  // ✅ FIX: response is ARRAY, not items
  const users = Array.isArray(res.data) ? res.data : [];

  console.log(`👥 Users fetched: ${users.length}`);

  const map = {};
  users.forEach(u => {
    map[u.userId] =
      `${u.firstName || ""} ${u.lastName || ""}`.trim() || u.email;
  });

  return map;
}

// ==============================
// HELPERS
// ==============================
function cleanHTML(text) {
  if (!text) return "";
  return text.replace(/<[^>]*>/g, "").trim();
}

function getDepartment(doc) {
  return doc.additionalInformation?.find(
    f => f.additionalFieldName === "Department"
  )?.value;
}

function getRequester(doc, userMap) {
  const event = doc.events?.find(e => e.eventType === "requesterSubmitted");
  return userMap[event?.authorId] || event?.authorId || "";
}

function getComments(doc, userMap) {
  return doc.events
    ?.filter(e => e.eventType === "comment")
    .map(e => `${userMap[e.authorId] || e.authorId}: ${cleanHTML(e.comment)}`)
    .join(" | ");
}

// ==============================
// GRAPH TOKEN
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
// RESET EXCEL (RECREATE TABLE)
// ==============================
async function resetExcel(token) {
  console.log("🧹 Resetting Excel...");

  const base =
    `https://graph.microsoft.com/v1.0/users/${process.env.EXCEL_USER}/drive/root:${process.env.EXCEL_FILE_PATH}`;

  // Clear sheet
  await axios.post(
    `${base}:/workbook/worksheets('${process.env.EXCEL_SHEET_NAME}')/range(address='A1:Z1000')/clear`,
    {},
    { headers: { Authorization: `Bearer ${token}` } }
  );

  // Add headers
  const headers = [[
    "requestId",
    "documentName",
    "description",
    "workflow",
    "status",
    "department",
    "requester",
    "comments",
    "createdAt",
    "modifiedAt",
    "decisionDate"
  ]];

  await axios.patch(
    `${base}:/workbook/worksheets('${process.env.EXCEL_SHEET_NAME}')/range(address='A1:K1')`,
    { values: headers },
    { headers: { Authorization: `Bearer ${token}` } }
  );

  // Create table
  await axios.post(
    `${base}:/workbook/tables/add`,
    {
      address: `${process.env.EXCEL_SHEET_NAME}!A1:K1`,
      hasHeaders: true
    },
    { headers: { Authorization: `Bearer ${token}` } }
  );

  console.log("✅ Excel ready");
}

// ==============================
// ADD ROWS
// ==============================
async function addRows(token, rows) {
  const url =
    `https://graph.microsoft.com/v1.0/users/${process.env.EXCEL_USER}/drive/root:${process.env.EXCEL_FILE_PATH}:/workbook/tables('${process.env.EXCEL_TABLE_NAME}')/rows/add`;

  await axios.post(url, { values: rows }, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    }
  });

  console.log(`✅ Inserted ${rows.length} rows`);
}

// ==============================
// MAIN
// ==============================
async function main() {
  try {
    console.log("🚀 Starting sync...");

    let token = (await getStoredTokens()).access_token;

    let docs;
    try {
      docs = await fetchAllDocuments(token);
    } catch (err) {
      if (err.response?.status === 401) {
        token = await refreshToken();
        docs = await fetchAllDocuments(token);
      } else throw err;
    }

    const userMap = await fetchUsers(token);

    const salesDocs = docs.filter(
      d => getDepartment(d) === "Sales & Pre-Sales"
    );

    console.log(`📊 Sales records: ${salesDocs.length}`);

    const rows = salesDocs.map(doc => [
      doc.requestId,
      doc.documentName,
      cleanHTML(doc.description),
      doc.friendlyName,
      doc.requestStatus,
      getDepartment(doc),
      getRequester(doc, userMap),
      getComments(doc, userMap),
      doc.createdAt,
      doc.modifiedAt,
      doc.decisionDate
    ]);

    const graphToken = await getGraphToken();

    await resetExcel(graphToken);
    await addRows(graphToken, rows);

    console.log("✅ SYNC COMPLETE");

  } catch (err) {
    console.error("❌ ERROR:", err.response?.data || err.message);
  }
}

main();