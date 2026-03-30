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
// FETCH DOCUMENTS
// ==============================
async function fetchAllDocuments(token) {
  let allDocs = [];
  let continuationToken = null;

  console.log("📄 Fetching documents...");

  while (true) {
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
    const batch = data.payload || [];
    console.log(`📄 Batch fetched: ${batch.length}`);
    allDocs.push(...batch);

    if (!data.continuationToken) {
      console.log("🛑 No continuation token → done");
      break;
    }

    continuationToken = data.continuationToken;
  }

  console.log(`✅ TOTAL DOCUMENTS: ${allDocs.length}`);
  return allDocs;
}

// ==============================
// FETCH USERS
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

// UPDATE 3: Format ISO timestamps to "2026-02-25, 7:44pm"
function formatDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  const date = d.toISOString().split("T")[0];
  let hours = d.getUTCHours();
  const minutes = d.getUTCMinutes().toString().padStart(2, "0");
  const ampm = hours >= 12 ? "pm" : "am";
  hours = hours % 12 || 12;
  return `${date}, ${hours}:${minutes}${ampm}`;
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
// CLEAR TABLE ROWS
// ==============================
async function clearTableRows(token) {
  console.log("🧹 Deleting table data rows...");

  const base =
    `https://graph.microsoft.com/v1.0/users/${process.env.EXCEL_USER}/drive/root:${process.env.EXCEL_FILE_PATH}`;

  let rangeRes;
  try {
    rangeRes = await axios.get(
      `${base}:/workbook/tables('${process.env.EXCEL_TABLE_NAME}')/dataBodyRange`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
  } catch (err) {
    if (
      err.response?.status === 404 ||
      err.response?.data?.error?.code === "ItemNotFound"
    ) {
      console.log("⚠️ Table is already empty, nothing to delete");
      return;
    }
    throw err;
  }

  const address = rangeRes.data.address;
  const [sheetName, rangePart] = address.includes("!")
    ? address.split("!")
    : [process.env.EXCEL_SHEET_NAME, address];

  console.log(`🗑️ Deleting rows: ${rangePart} on sheet: "${sheetName}"`);

  await axios.post(
    `${base}:/workbook/worksheets('${encodeURIComponent(sheetName)}')/range(address='${rangePart}')/delete`,
    { shift: "Up" },
    {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      }
    }
  );

  console.log("✅ Table rows deleted");
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

    // UPDATE 6: Filter by "Refund Request" workflow only
    // (replaces the department filter — all refund requests are Sales & Pre-Sales)
    const refundDocs = docs.filter(d =>
      d.friendlyName?.startsWith("Refund Request") &&
      d.documentName?.trim() &&
      cleanHTML(d.description)?.trim()
    );

    console.log(`📊 Refund Request records: ${refundDocs.length}`);

    // Columns: Request ID | Document Name | Description | Status |
    //          Requester | Comments | Created At | Decision Date
    const rows = refundDocs.map(doc => [
      doc.requestId,                          // UPDATE 4: department removed
      doc.documentName,                       // UPDATE 5: friendlyName/workflow removed
      cleanHTML(doc.description),
      doc.requestStatus,
      getRequester(doc, userMap),
      getComments(doc, userMap),
      formatDate(doc.createdAt),              // UPDATE 3: formatted
      formatDate(doc.decisionDate)            // UPDATE 2: modifiedAt removed, UPDATE 3: formatted
    ]);

    const graphToken = await getGraphToken();
    await clearTableRows(graphToken);
    await addRows(graphToken, rows);

    console.log("✅ SYNC COMPLETE");

  } catch (err) {
    console.error("❌ ERROR:", err.response?.data || err.message);
  }
}

// ==============================
// CRON SAFE EXIT
// ==============================
main()
  .then(() => process.exit(0))
  .catch(() => process.exit(1));