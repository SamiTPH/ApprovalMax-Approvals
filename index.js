require("dotenv").config();
const axios = require("axios");
const { Pool } = require("pg");

// ==============================
// DB
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
// FETCH DOCUMENTS (DATE FILTER)
// ==============================
async function fetchDocuments() {
  let token = (await getStoredTokens()).access_token;

  try {
    console.log("📅 Fetching last 90 days of documents...");

    const fromDate = new Date();
    fromDate.setDate(fromDate.getDate() - 90);

    const res = await axios.get(
      `https://public-api.approvalmax.com/api/v1/companies/${process.env.COMPANY_ID}/standalone/documents?createdFrom=${fromDate.toISOString()}`,
      {
        headers: { Authorization: `Bearer ${token}` }
      }
    );

    const docs = res.data.payload || [];

    console.log(`📄 Documents fetched: ${docs.length}`);

    return docs;

  } catch (err) {
    if (err.response?.status === 401) {
      console.log("⚠️ Token expired → refreshing...");
      await refreshToken();
      return fetchDocuments();
    }

    throw err;
  }
}

// ==============================
// FETCH USERS (WITH RETRY)
// ==============================
async function fetchUsers() {
  let token = (await getStoredTokens()).access_token;

  try {
    console.log("👤 Fetching users...");

    const res = await axios.get(
      `https://public-api.approvalmax.com/api/v1/companies/${process.env.COMPANY_ID}/userProfiles`,
      {
        headers: { Authorization: `Bearer ${token}` }
      }
    );

    const users = res.data.payload || [];

    console.log(`👥 Users fetched: ${users.length}`);

    const map = {};
    users.forEach(u => {
      map[u.userId] = `${u.firstName} ${u.lastName}`;
    });

    return map;

  } catch (err) {
    if (err.response?.status === 401) {
      console.log("⚠️ Token expired (users) → refreshing...");
      await refreshToken();
      return fetchUsers();
    }

    throw err;
  }
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
  console.log("🔐 Getting Graph token...");

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
// CLEAR EXCEL (CORRECT METHOD)
// ==============================
async function clearExcel(token) {
  console.log("🧹 Clearing Excel (range.clear)...");

  const url =
    `https://graph.microsoft.com/v1.0/users/${process.env.EXCEL_USER}/drive/root:${process.env.EXCEL_FILE_PATH}:/workbook/tables/${process.env.EXCEL_TABLE_NAME}/range/clear`;

  await axios.post(url, {}, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    }
  });

  console.log("✅ Excel cleared");
}

// ==============================
// ADD ROWS
// ==============================
async function addRows(token, rows) {
  if (!rows.length) {
    console.log("⚠️ No rows to insert");
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

  console.log(`✅ Inserted ${rows.length} rows`);
}

// ==============================
// MAIN
// ==============================
async function main() {
  try {
    console.log("🚀 Starting sync...");

    const docs = await fetchDocuments();
    const userMap = await fetchUsers();

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

    await clearExcel(graphToken);
    await addRows(graphToken, rows);

    console.log("✅ FULL SYNC COMPLETE");

  } catch (err) {
    console.error("❌ ERROR:", err.response?.data || err.message);
  }
}

main();