import { mainSync } from "../index.js";

export default async function handler(req, res) {
  // Log every request
  const timestamp = new Date().toISOString();
  console.log(`\n${'='.repeat(60)}`);
  console.log(`🔔 SYNC ENDPOINT CALLED`);
  console.log(`   Time: ${timestamp}`);
  console.log(`   Method: ${req.method}`);
  console.log(`   URL: ${req.url}`);
  console.log(`${'='.repeat(60)}\n`);

  if (req.method !== "GET") {
    console.log("❌ Wrong method - returning 405");
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    // Check environment variables
    console.log("🔍 Checking environment variables...");
    const envCheck = {
      GOOGLE_SHEET_ID: !!process.env.GOOGLE_SHEET_ID,
      GOOGLE_CLIENT_EMAIL: !!process.env.GOOGLE_CLIENT_EMAIL,
      GOOGLE_PRIVATE_KEY: !!process.env.GOOGLE_PRIVATE_KEY,
      GOOGLE_PROJECT_ID: !!process.env.GOOGLE_PROJECT_ID,
      EMAIL_PROVIDER: process.env.EMAIL_PROVIDER || "NOT SET",
      EMAIL_USER: !!process.env.EMAIL_USER,
      EMAIL_PASS: !!process.env.EMAIL_PASS,
      SENDGRID_API_KEY: !!process.env.SENDGRID_API_KEY,
    };
    
    console.log("Environment variables status:");
    Object.entries(envCheck).forEach(([key, value]) => {
      const status = typeof value === 'boolean' ? (value ? '✅' : '❌') : `📝 ${value}`;
      console.log(`   ${key}: ${status}`);
    });

    // Check for missing critical env vars
    const missingVars = [];
    if (!process.env.GOOGLE_SHEET_ID) missingVars.push("GOOGLE_SHEET_ID");
    if (!process.env.GOOGLE_CLIENT_EMAIL) missingVars.push("GOOGLE_CLIENT_EMAIL");
    if (!process.env.GOOGLE_PRIVATE_KEY) missingVars.push("GOOGLE_PRIVATE_KEY");
    if (!process.env.EMAIL_PROVIDER) missingVars.push("EMAIL_PROVIDER");
    
    if (missingVars.length > 0) {
      const errorMsg = `Missing required environment variables: ${missingVars.join(", ")}`;
      console.error(`❌ ${errorMsg}`);
      return res.status(500).json({ 
        ok: false, 
        error: errorMsg,
        envCheck 
      });
    }

    console.log("\n🚀 Starting mainSync()...\n");
    
    const result = await mainSync();
    
    console.log("\n✅ mainSync() completed successfully!");
    console.log("Result:", JSON.stringify(result, null, 2));
    
    const response = {
      ok: true,
      timestamp,
      ...result
    };
    
    console.log("\n📤 Sending response:", JSON.stringify(response, null, 2));
    res.status(200).json(response);
    
  } catch (e) {
    console.error("\n❌ ERROR OCCURRED:");
    console.error("Message:", e.message);
    console.error("Stack:", e.stack);
    
    const errorResponse = {
      ok: false,
      timestamp,
      error: e.message,
      stack: e.stack
    };
    
    console.log("\n📤 Sending error response:", JSON.stringify(errorResponse, null, 2));
    res.status(500).json(errorResponse);
  }
}