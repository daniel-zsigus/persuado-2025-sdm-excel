const express = require("express");
const fetch = require("node-fetch");
const app = express();
app.use(express.json());

const CLIENT_ID = "EvglXZt0zxqDO4pNdRPxrPBiHESGWDqn";
const CLIENT_SECRET = "ATOAHQqcpoRMY3C8DSCYJJz5WoqzK2aCyLMk9JdapCxQYhOzBOTNSy0GGYJTF2WRvl2rDC60E23A";
const REDIRECT_URI = "http://localhost:3000/callback"; 

// Token exchange endpoint
app.post("/exchange-token", async (req, res) => {
  const { code } = req.body;

  try {
    const response = await fetch("https://auth.atlassian.com/oauth/token", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        grant_type: "authorization_code",
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        code,
        redirect_uri: REDIRECT_URI,
      }),
    });

    const data = await response.json();
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Token exchange failed" });
  }
});

// Callback page
app.get("/callback", (req, res) => {
  const { code } = req.query;

  res.send(`
    <html>
      <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
      <script>
        Office.onReady(() => {
          Office.context.ui.messageParent(JSON.stringify({ status: "ok", code: "${code}" }));
        });
      </script>
      <body>You are logged in! You can close this window.</body>
    </html>
  `);
});

app.listen(3000, () => console.log("✅ Server running on http://localhost:3000"));
