/**
 * ============================================================================
 * EXPIRYGO BASELINE LOAD TESTING ENGINE
 * ============================================================================
 * Simulates normal expected concurrent traffic:
 *  - 100 Virtual Users (VUs)
 *  - Running continuously for 1 Minute (60 seconds)
 *  - Thousands of requests across representative API endpoints
 *  - Real-time RPS & Latency Metrics: Min, Avg, Max, p50, p90, p95, p99
 *  - Generates comprehensive Excel and Markdown reports
 * ============================================================================
 */

import http from "http";
import https from "https";
import { performance } from "perf_hooks";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { execSync } from "child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Ensure results directory exists
const RESULTS_DIR = path.join(__dirname, "results");
if (!fs.existsSync(RESULTS_DIR)) {
  fs.mkdirSync(RESULTS_DIR, { recursive: true });
}

// ----------------------------------------------------------------------------
// CLI Arguments Parsing
// ----------------------------------------------------------------------------
const args = process.argv.slice(2);
function getArg(flag, defaultValue) {
  const index = args.indexOf(flag);
  if (index !== -1 && args[index + 1]) {
    return args[index + 1];
  }
  return defaultValue;
}

const BASE_URL = getArg("--url", "http://127.0.0.1:8000");
const VIRTUAL_USERS = parseInt(getArg("--users", "100"), 10);
const DURATION_SECONDS = parseInt(getArg("--duration", "60"), 10);
const RAMP_UP_SECONDS = parseInt(getArg("--ramp", "3"), 10);

console.log("\n" + "=".repeat(78));
console.log("🚀 EXPIRYGO BASELINE LOAD TESTING HARNESS");
console.log("=".repeat(78));
console.log(` Target Host       : ${BASE_URL}`);
console.log(` Virtual Users     : ${VIRTUAL_USERS} concurrent users`);
console.log(` Test Duration     : ${DURATION_SECONDS} seconds (1 minute sustained run)`);
console.log(` Ramp-Up Time      : ${RAMP_UP_SECONDS} seconds`);
console.log(` Results Directory : ${RESULTS_DIR}`);
console.log("=".repeat(78) + "\n");

// ----------------------------------------------------------------------------
// Workload Endpoints Definition (Realistic Customer Browsing Pattern)
// ----------------------------------------------------------------------------
const ENDPOINTS = [
  { path: "/health", weight: 20, name: "Health Check (Gateway)" },
  { path: "/health/db", weight: 15, name: "Database Health Probe" },
  { path: "/health/platform-impact", weight: 15, name: "Platform Impact Metrics" },
  { path: "/products/", weight: 25, name: "Browse Products Catalog" },
  { path: "/shops/", weight: 15, name: "List Verified Shops" },
  { path: "/products/search/deep?q=milk", weight: 5, name: "Deep Semantic Search" },
  { path: "/translate/languages", weight: 5, name: "Multilingual UI Locales" },
];

// Weighted random endpoint selection
function getRandomEndpoint() {
  const totalWeight = ENDPOINTS.reduce((sum, ep) => sum + ep.weight, 0);
  let random = Math.random() * totalWeight;
  for (const ep of ENDPOINTS) {
    if (random < ep.weight) return ep;
    random -= ep.weight;
  }
  return ENDPOINTS[0];
}

// High performance HTTP Agent with connection reuse
const parsedUrl = new URL(BASE_URL);
const isHttps = parsedUrl.protocol === "https:";
const agentOptions = {
  keepAlive: true,
  maxSockets: VIRTUAL_USERS * 2,
  maxFreeSockets: 50,
  timeout: 10000,
};
const httpAgent = isHttps ? new https.Agent(agentOptions) : new http.Agent(agentOptions);

// ----------------------------------------------------------------------------
// Metrics Storage
// ----------------------------------------------------------------------------
let totalRequests = 0;
let totalSuccess = 0;
let totalErrors = 0;
const statusCodes = {};
const allLatencies = [];
const endpointStats = {};
const secondBuckets = [];

ENDPOINTS.forEach((ep) => {
  endpointStats[ep.path] = {
    name: ep.name,
    count: 0,
    success: 0,
    errors: 0,
    latencies: [],
    min: Infinity,
    max: 0,
    sum: 0,
  };
});

// Helper: Make a single HTTP request and record high-res timing
function makeRequest(endpointObj) {
  return new Promise((resolve) => {
    const startTime = performance.now();
    const requestOptions = {
      hostname: parsedUrl.hostname,
      port: parsedUrl.port || (isHttps ? 443 : 80),
      path: endpointObj.path,
      method: "GET",
      agent: httpAgent,
      headers: {
        "User-Agent": "ExpiryGo-LoadTest-VU/1.0",
        Accept: "application/json",
        Connection: "keep-alive",
      },
    };

    const client = isHttps ? https : http;
    const req = client.request(requestOptions, (res) => {
      let data = "";
      res.on("data", (chunk) => {
        data += chunk;
      });
      res.on("end", () => {
        const endTime = performance.now();
        const latency = endTime - startTime;
        resolve({
          statusCode: res.statusCode,
          latency,
          error: null,
          endpoint: endpointObj.path,
        });
      });
    });

    req.on("error", (err) => {
      const endTime = performance.now();
      resolve({
        statusCode: 0,
        latency: endTime - startTime,
        error: err.message,
        endpoint: endpointObj.path,
      });
    });

    req.setTimeout(10000, () => {
      req.destroy(new Error("Request Timeout (10s)"));
    });

    req.end();
  });
}

// ----------------------------------------------------------------------------
// Virtual User (Worker) Coroutine
// ----------------------------------------------------------------------------
async function virtualUser(userId, stopTime) {
  // Staggered ramp-up delay
  const rampDelay = (RAMP_UP_SECONDS * 1000 * userId) / VIRTUAL_USERS;
  await new Promise((r) => setTimeout(r, rampDelay));

  while (Date.now() < stopTime) {
    const endpoint = getRandomEndpoint();
    const result = await makeRequest(endpoint);

    // Record metrics
    totalRequests++;
    const code = result.statusCode || "ERR";
    statusCodes[code] = (statusCodes[code] || 0) + 1;

    const lat = result.latency;
    allLatencies.push(lat);

    const epStat = endpointStats[endpoint.path];
    if (epStat) {
      epStat.count++;
      epStat.sum += lat;
      epStat.latencies.push(lat);
      if (lat < epStat.min) epStat.min = lat;
      if (lat > epStat.max) epStat.max = lat;

      if (result.statusCode >= 200 && result.statusCode < 400) {
        epStat.success++;
        totalSuccess++;
      } else {
        epStat.errors++;
        totalErrors++;
      }
    }

    // Realistic micro think-time between calls (5ms - 25ms)
    const thinkTime = Math.floor(Math.random() * 20) + 5;
    await new Promise((r) => setTimeout(r, thinkTime));
  }
}

// ----------------------------------------------------------------------------
// Statistical Calculations
// ----------------------------------------------------------------------------
function calculatePercentile(sortedArray, percentile) {
  if (sortedArray.length === 0) return 0;
  const index = Math.ceil((percentile / 100) * sortedArray.length) - 1;
  return sortedArray[Math.max(0, Math.min(index, sortedArray.length - 1))];
}

// ----------------------------------------------------------------------------
// Main Execution Loop
// ----------------------------------------------------------------------------
async function runLoadTest() {
  const startTime = Date.now();
  const stopTime = startTime + DURATION_SECONDS * 1000;

  console.log("⏳ Initializing 100 Virtual Users...");

  // Launch virtual user coroutines
  const userPromises = [];
  for (let i = 0; i < VIRTUAL_USERS; i++) {
    userPromises.push(virtualUser(i, stopTime));
  }

  // Ticker for real-time console telemetry
  let lastRequestCount = 0;
  let secondCounter = 0;

  const tickerInterval = setInterval(() => {
    secondCounter++;
    const currentTotal = totalRequests;
    const requestsThisSecond = currentTotal - lastRequestCount;
    lastRequestCount = currentTotal;

    const elapsed = Math.floor((Date.now() - startTime) / 1000);
    const recentLatencies = allLatencies.slice(-Math.max(requestsThisSecond, 1));
    const recentAvg =
      recentLatencies.length > 0
        ? recentLatencies.reduce((a, b) => a + b, 0) / recentLatencies.length
        : 0;
    const recentMin = recentLatencies.length > 0 ? Math.min(...recentLatencies) : 0;
    const recentMax = recentLatencies.length > 0 ? Math.max(...recentLatencies) : 0;

    // Store second bucket
    secondBuckets.push({
      second: secondCounter,
      rps: requestsThisSecond,
      avgLatency: parseFloat(recentAvg.toFixed(2)),
      minLatency: parseFloat(recentMin.toFixed(2)),
      maxLatency: parseFloat(recentMax.toFixed(2)),
      totalRequestsSoFar: currentTotal,
      errorsSoFar: totalErrors,
    });

    const elapsedFormatted = `${String(Math.floor(elapsed / 60)).padStart(2, "0")}:${String(
      elapsed % 60
    ).padStart(2, "0")}`;
    const totalFormatted = `${String(Math.floor(DURATION_SECONDS / 60)).padStart(2, "0")}:${String(
      DURATION_SECONDS % 60
    ).padStart(2, "0")}`;

    process.stdout.write(
      `\r⏱️  [${elapsedFormatted} / ${totalFormatted}] VUs: ${VIRTUAL_USERS} | ` +
        `Current: \x1b[32m${requestsThisSecond} req/sec\x1b[0m | ` +
        `Total: \x1b[36m${currentTotal.toLocaleString()}\x1b[0m | ` +
        `Avg: \x1b[33m${recentAvg.toFixed(1)}ms\x1b[0m | ` +
        `Min: ${recentMin.toFixed(1)}ms | ` +
        `Max: ${recentMax.toFixed(1)}ms | ` +
        `Errors: ${totalErrors > 0 ? "\x1b[31m" + totalErrors + "\x1b[0m" : "0"}`
    );
  }, 1000);

  // Wait for all workers to complete
  await Promise.all(userPromises);
  clearInterval(tickerInterval);
  console.log("\n\n✅ Load test completed successfully!\n");

  // --------------------------------------------------------------------------
  // Compute Final Summary Metrics
  // --------------------------------------------------------------------------
  const totalDurationActual = (Date.now() - startTime) / 1000;
  const sortedLatencies = [...allLatencies].sort((a, b) => a - b);
  const minLatency = sortedLatencies.length > 0 ? sortedLatencies[0] : 0;
  const maxLatency = sortedLatencies.length > 0 ? sortedLatencies[sortedLatencies.length - 1] : 0;
  const sumLatency = sortedLatencies.reduce((a, b) => a + b, 0);
  const avgLatency = sortedLatencies.length > 0 ? sumLatency / sortedLatencies.length : 0;

  const p50 = calculatePercentile(sortedLatencies, 50);
  const p75 = calculatePercentile(sortedLatencies, 75);
  const p90 = calculatePercentile(sortedLatencies, 90);
  const p95 = calculatePercentile(sortedLatencies, 95);
  const p99 = calculatePercentile(sortedLatencies, 99);
  const overallRPS = totalRequests / totalDurationActual;
  const successRate = totalRequests > 0 ? (totalSuccess / totalRequests) * 100 : 0;

  // Print Formatted Report Table
  console.log("=".repeat(78));
  console.log("📊 EXPIRYGO BASELINE LOAD TEST RESULTS SUMMARY");
  console.log("=".repeat(78));
  console.log(` Concurrent Virtual Users : ${VIRTUAL_USERS}`);
  console.log(` Sustained Duration       : ${totalDurationActual.toFixed(1)} seconds`);
  console.log(` Total Requests Executed  : ${totalRequests.toLocaleString()}`);
  console.log(` Requests Per Second (RPS): \x1b[1;32m${overallRPS.toFixed(2)} req/sec\x1b[0m`);
  console.log(` Success Rate             : \x1b[1;32m${successRate.toFixed(2)}%\x1b[0m`);
  console.log("-".repeat(78));
  console.log("⏱️  RESPONSE TIME PROFILE:");
  console.log(`   • Minimum Response Time : ${minLatency.toFixed(2)} ms`);
  console.log(`   • Average Response Time : \x1b[1;33m${avgLatency.toFixed(2)} ms\x1b[0m`);
  console.log(`   • Median (p50)          : ${p50.toFixed(2)} ms`);
  console.log(`   • 90th Percentile (p90) : ${p90.toFixed(2)} ms`);
  console.log(`   • 95th Percentile (p95) : ${p95.toFixed(2)} ms`);
  console.log(`   • 99th Percentile (p99) : ${p99.toFixed(2)} ms`);
  console.log(`   • Maximum Response Time : ${maxLatency.toFixed(2)} ms (${(maxLatency / 1000).toFixed(2)}s)`);
  console.log("-".repeat(78));
  console.log("📡 HTTP STATUS CODES:");
  Object.keys(statusCodes).forEach((code) => {
    console.log(`   • HTTP ${code} : ${statusCodes[code].toLocaleString()} responses`);
  });
  console.log("=".repeat(78));

  // Endpoint Details Table
  console.log("\n📍 ENDPOINT BREAKDOWN:");
  console.log(
    "┌──────────────────────────────┬──────────┬──────────┬──────────┬──────────┬──────────┐"
  );
  console.log(
    "│ Endpoint                     │ Requests │ RPS      │ Avg (ms) │ Min (ms) │ Max (ms) │"
  );
  console.log(
    "├──────────────────────────────┼──────────┼──────────┼──────────┼──────────┼──────────┤"
  );

  const endpointSummaryArray = [];
  ENDPOINTS.forEach((ep) => {
    const s = endpointStats[ep.path];
    const epRPS = s.count / totalDurationActual;
    const epAvg = s.count > 0 ? s.sum / s.count : 0;
    const epMin = s.min === Infinity ? 0 : s.min;
    const epMax = s.max;
    const epSorted = [...s.latencies].sort((a, b) => a - b);
    const epP95 = calculatePercentile(epSorted, 95);

    endpointSummaryArray.push({
      path: ep.path,
      name: s.name,
      requests: s.count,
      rps: parseFloat(epRPS.toFixed(2)),
      avgMs: parseFloat(epAvg.toFixed(2)),
      minMs: parseFloat(epMin.toFixed(2)),
      maxMs: parseFloat(epMax.toFixed(2)),
      p95Ms: parseFloat(epP95.toFixed(2)),
      successCount: s.success,
      errorCount: s.errors,
    });

    const pathPadded = (ep.path.length > 28 ? ep.path.slice(0, 25) + "..." : ep.path).padEnd(28);
    const reqPadded = String(s.count).padStart(8);
    const rpsPadded = epRPS.toFixed(1).padStart(8);
    const avgPadded = epAvg.toFixed(1).padStart(8);
    const minPadded = epMin.toFixed(1).padStart(8);
    const maxPadded = epMax.toFixed(1).padStart(8);

    console.log(`│ ${pathPadded} │ ${reqPadded} │ ${rpsPadded} │ ${avgPadded} │ ${minPadded} │ ${maxPadded} │`);
  });
  console.log(
    "└──────────────────────────────┴──────────┴──────────┴──────────┴──────────┴──────────┘\n"
  );

  // Save raw results to JSON
  const rawResults = {
    meta: {
      baseUrl: BASE_URL,
      virtualUsers: VIRTUAL_USERS,
      configuredDuration: DURATION_SECONDS,
      actualDuration: parseFloat(totalDurationActual.toFixed(2)),
      timestamp: new Date().toISOString(),
    },
    summary: {
      totalRequests,
      totalSuccess,
      totalErrors,
      requestsPerSecond: parseFloat(overallRPS.toFixed(2)),
      successRatePercentage: parseFloat(successRate.toFixed(2)),
      latencyMs: {
        min: parseFloat(minLatency.toFixed(2)),
        avg: parseFloat(avgLatency.toFixed(2)),
        p50: parseFloat(p50.toFixed(2)),
        p75: parseFloat(p75.toFixed(2)),
        p90: parseFloat(p90.toFixed(2)),
        p95: parseFloat(p95.toFixed(2)),
        p99: parseFloat(p99.toFixed(2)),
        max: parseFloat(maxLatency.toFixed(2)),
      },
      statusCodes,
    },
    endpoints: endpointSummaryArray,
    timeSeries: secondBuckets,
  };

  const rawJsonPath = path.join(RESULTS_DIR, "baseline-load-test-raw.json");
  fs.writeFileSync(rawJsonPath, JSON.stringify(rawResults, null, 2), "utf-8");
  console.log(`💾 Raw test metrics saved to: ${rawJsonPath}`);

  // Auto-generate Excel & Markdown Reports
  console.log("📊 Triggering automated Excel and Markdown report generation...");
  try {
    execSync("node generate-load-test-excel.js", {
      cwd: __dirname,
      stdio: "inherit",
    });
  } catch (err) {
    console.error("Error generating Excel report:", err.message);
  }
}

runLoadTest().catch((err) => {
  console.error("Fatal Load Test Error:", err);
  process.exit(1);
});
