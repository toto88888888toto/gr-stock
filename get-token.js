/**
 * Run: GOOGLE_CLIENT_ID=xxx GOOGLE_CLIENT_SECRET=yyy node get-token.js
 * Or set them in your shell first, then: node get-token.js
 */
const { google } = require("googleapis");
const http = require("http");
const url = require("url");

const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT = "http://localhost:3333/callback";

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error("Set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET env vars first");
  process.exit(1);
}

const oauth2 = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT);

const authUrl = oauth2.generateAuthUrl({
  access_type: "offline",
  prompt: "consent",
  scope: ["https://www.googleapis.com/auth/drive"],
});

console.log("\n✅ Open this URL in your browser:\n");
console.log(authUrl);
console.log("\nWaiting for Google to redirect...\n");

const server = http.createServer(async (req, res) => {
  const code = new url.URL(req.url, "http://localhost:3333").searchParams.get("code");
  if (!code) { res.end("No code"); return; }
  try {
    const { tokens } = await oauth2.getToken(code);
    res.end("<h2>✅ Success! Check your terminal.</h2>");
    console.log("\n🎉 REFRESH TOKEN:\n");
    console.log(tokens.refresh_token);
  } catch (e) {
    res.end("Error: " + e.message);
    console.error("Error:", e.message);
  }
  server.close();
});

server.listen(3333);
