# 🚀 ExpiryGo Baseline & Stress Load Testing Suite

Comprehensive load testing suite designed to validate API throughput, concurrency capacity, and response time SLAs under sustained normal traffic.

---

## 📋 Test Specification

* **Virtual Users (VUs):** `100` concurrent virtual shoppers
* **Execution Duration:** `60 seconds` (1 minute sustained run)
* **Expected Throughput:** Thousands of continuous requests (~100–200 req/sec)
* **Workload Mix:** Realistic customer browsing patterns:
  - `GET /health` (Gateway liveness)
  - `GET /products/` (Catalog listings)
  - `GET /shops/` (Verified stores directory)
  - `GET /products/categories` (Categories)
  - `GET /products/flash-deals` (Expiring surplus deals)
  - `GET /products/surprise-bags` (Mystery food bags)
  - `GET /shops/map` (Geospatial store coordinates)
  - `GET /products/search/deep?q=bakery` (Deep semantic search)

---

## 📊 Metrics Captured & Reported

1. **Requests Per Second (RPS):**
   * Instantaneous second-by-second throughput.
   * Average sustained RPS throughout the 1-minute window.
2. **Response Time Profile:**
   * **Minimum (Fastest response):** Sub-millisecond/in-memory timing.
   * **Average:** Mean response latency across all endpoints.
   * **Maximum (Slowest response):** Worst-case execution ceiling.
   * **Percentiles:** Median (p50), p75, p90, p95, p99.
3. **HTTP Status Code Breakdown:**
   * 200 OK, 3xx redirects, 429 Rate Limited, 5xx Server Errors.
4. **Endpoint Performance Matrix:**
   * Route-by-route RPS, latency distribution, and error count.

---

## 🛠️ How to Run

### 1. Ensure Backend Server is Running
In a terminal:
```bash
cd backend
python -m uvicorn main:app --host 127.0.0.1 --port 8000
```

### 2. Run the Full 1-Minute Baseline Load Test
```bash
cd load-tests
npm run test:full
```
Or directly:
```bash
node run-baseline-load-test.js --users 100 --duration 60
```

### 3. Run a Quick 15-Second Verification Test
```bash
cd load-tests
npm run test:quick
```

### 4. Regenerate Excel & Markdown Reports
```bash
cd load-tests
npm run report:excel
```

---

## 📁 Generated Deliverables

Results are automatically saved to `load-tests/results/`:

* 📊 **`ExpiryGo_Baseline_Load_Test_Report.xlsx`** (Multi-sheet Excel workbook):
  - **Sheet 1:** Executive Dashboard & KPIs
  - **Sheet 2:** Endpoint Performance Breakdown
  - **Sheet 3:** Second-by-Second Time Series Log (60 data points)
  - **Sheet 4:** Latency Percentiles & SLAs (p50, p75, p90, p95, p99, Min, Max)
  - **Sheet 5:** Capacity & High-Concurrency Tuning Guide
* 📄 **`baseline-load-test-report.md`** — Markdown documentation.
* 💾 **`baseline-load-test-raw.json`** — Raw telemetry and time series metrics.
