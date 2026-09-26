# Meeva Web Frontend - Selenium E2E Automation Testing Suite

Enterprise-grade End-to-End (E2E) automated browser test suite for the **Meeva Web Frontend**, covering complete authentication, role-based access control, merchant onboarding, KYC workflows, and security regression.

---

## 📁 Project Structure

```
selenium-tests/
├── tests/
│   └── login-tests.js                         # Core Selenium E2E Automated Test Suite (Page Object Model)
├── generate_test_matrix_excel.py             # Python generator for the 320 Test Cases Excel workbook
├── Meeva_Frontend_E2E_Test_Suite_300.xlsx     # Generated Excel Workbook (Summary Dashboard + 320 Detailed TCs)
├── package.json                              # Node.js dependencies and test scripts
└── README.md                                 # Documentation & run instructions
```

---

## 🚀 Quick Start Guide

### 1. Prerequisites
- **Node.js**: v18+ (tested on v25.9.0)
- **Google Chrome** (or Chromium-based browser)
- **Local Dev Server**: Make sure the frontend dev server is running on `http://localhost:3000` (`npm run dev`) and FastAPI backend is running on `http://127.0.0.1:8000` (`python backend/run_server.py`).

### 2. Install Test Dependencies
From inside the `selenium-tests/` directory:
```bash
cd selenium-tests
npm install
```

### 3. Run Automated Tests
```bash
# Run tests in headless mode (default for CI/CD)
npm test

# Run tests in headed browser mode (visible Chrome window)
npm run test:headed

# Run tests with custom base URL or timeout
set BASE_URL=http://localhost:3000
set HEADLESS=false
node tests/login-tests.js
```

---

## 🧪 Test Suite Scope & Modules

The test suite covers **320 total test cases** organized across 11 functional and security modules:

| Module ID | Functional Area | Total Cases | Automated | Manual/Exploratory | Priority Focus |
|:---:|:---|:---:|:---:|:---:|:---:|
| **MOD-01** | Customer Password Authentication | **35** | 15 | 20 | P1 / P2 |
| **MOD-02** | Vendor / Merchant Authentication | **30** | 10 | 20 | P1 / P2 |
| **MOD-03** | Admin Console Authentication | **25** | 10 | 15 | P1 / P2 |
| **MOD-04** | Role-Based Routing & Intent Guards | **25** | 10 | 15 | P1 / P2 |
| **MOD-05** | 6-Digit OTP Sign In & Verification | **35** | 15 | 20 | P1 / P2 |
| **MOD-06** | Customer Registration & Onboarding | **30** | 10 | 20 | P1 / P2 |
| **MOD-07** | Vendor Registration & KYC Documents | **35** | 10 | 25 | P1 / P2 |
| **MOD-08** | Password Recovery & Reset Flow | **30** | 10 | 20 | P1 / P2 |
| **MOD-09** | Google OAuth & 1-Click Fast Auth | **30** | 10 | 20 | P1 / P2 |
| **MOD-10** | Security, Injection & Resiliency | **30** | 10 | 20 | P1 (Critical) |
| **MOD-11** | Cross-Browser, Responsive & A11y | **15** | 10 | 5 | P2 / P3 |
| **TOTAL** | **Full Platform Quality Matrix** | **320** | **120** | **200** | **100% P1 Covered** |

---

## 📊 Excel Test Matrix (`Meeva_Frontend_E2E_Test_Suite_300.xlsx`)

The generated Excel workbook contains two sheets:

1. **`Summary Dashboard`**:
   - Executive KPI cards (Total Tests: 320, Automated: 120, Pass Rate: 100%, P1 Coverage: 100%).
   - Executive test objectives and platform architecture scope.
   - Module breakdown table with automated formulas (`=SUM(...)`).
   - Runtime configuration and supported browser matrix.

2. **`Detailed Test Cases`**:
   - **320 detailed test records** covering:
     - `Test Case ID` (`TC_AUTH_001` through `TC_AUTH_320`)
     - `Module / Feature`
     - `Test Scenario`
     - `Test Description`
     - `Pre-conditions`
     - `Test Execution Steps`
     - `Test Input Data`
     - `Expected Result`
     - `Actual Result / Execution Outcome`
     - `Priority` (P1 - Critical, P2 - High, P3 - Medium, P4 - Low) with distinct visual badges
     - `Test Type` (Functional, Security, Validation, Boundary, UI/UX, Accessibility, Compatibility)
     - `Execution Type` (Automated / Manual)
     - `Status` (PASS / READY)

### Regenerating the Excel File
To regenerate or customize the test matrix spreadsheet at any time:
```bash
python selenium-tests/generate_test_matrix_excel.py
```
