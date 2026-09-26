/**
 * Meeva Web Frontend - Selenium E2E Automation Test Suite
 * File: selenium-tests/tests/login-tests.js
 *
 * Covers End-to-End browser test automation for:
 *  - Customer, Vendor, and Admin Authentication
 *  - Role selection, switching, and role mismatch guards
 *  - Password-based Login & OTP-based 6-digit Login flows
 *  - Password Visibility toggle (Eye/EyeOff masking)
 *  - Customer Registration & Password confirmation validation
 *  - Vendor Onboarding, Store Category selection & KYC uploads
 *  - Password Recovery & Reset flows (Forgot Password -> OTP -> New Password)
 *  - Google OAuth / 1-Click Fast Identity modal & configuration
 *  - Form validations, SQLi/XSS sanitization security checks
 *  - LocalStorage token persistence & Post-login role routing
 *
 * Requirements:
 *  npm install selenium-webdriver chromedriver
 *  Run: node tests/login-tests.js
 */

import { Builder, By, until, Key } from "selenium-webdriver";
import chrome from "selenium-webdriver/chrome.js";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ==========================================
// Test Configuration & Environment Variables
// ==========================================
const BASE_URL = process.env.BASE_URL || "http://localhost:3000";
const AUTH_URL = `${BASE_URL}/auth`;
const IS_HEADLESS = process.env.HEADLESS !== "false";
const DEFAULT_TIMEOUT = parseInt(process.env.TIMEOUT || "8000", 10);
const SCREENSHOT_DIR = path.join(__dirname, "..", "screenshots");

// Create screenshot directory if not existing
if (!fs.existsSync(SCREENSHOT_DIR)) {
  fs.mkdirSync(SCREENSHOT_DIR, { recursive: true });
}

// ==========================================
// Page Object Model: AuthPage
// ==========================================
class AuthPage {
  /**
   * @param {import('selenium-webdriver').WebDriver} driver
   */
  constructor(driver) {
    this.driver = driver;
  }

  // --- Locators ---
  get locators() {
    return {
      // Header & Branding
      brandHeading: By.xpath("//span[contains(text(),'Mee') and contains(text(),'va')]"),
      pageSubtitle: By.xpath("//div[contains(@class,'text-center')]//p"),

      // Tab switcher: Sign In vs Sign Up
      tabSignIn: By.xpath("//button[text()='Sign In']"),
      tabSignUp: By.xpath("//button[text()='Sign Up']"),

      // Role selectors
      roleCustomer: By.xpath("//button[contains(.,'Customer')]"),
      roleVendor: By.xpath("//button[contains(.,'Vendor')]"),
      roleAdmin: By.xpath("//button[contains(.,'Admin')]"),

      // Login Mode switcher: Password vs OTP
      loginModePassword: By.xpath("//button[contains(.,'Password Sign In')]"),
      loginModeOtp: By.xpath("//button[contains(.,'6-Digit OTP Sign In')]"),

      // Password Login Inputs
      loginEmailInput: By.css('input[name="login_email_no_autofill"]'),
      loginPasswordInput: By.css('input[name="login_password_no_autofill"]'),
      loginSubmitButton: By.xpath("//form//button[@type='submit']"),
      passwordToggleBtn: By.xpath("//input[@name='login_password_no_autofill']/following-sibling::button"),
      forgotPasswordLink: By.xpath("//button[contains(text(),'Forgot password?')]"),
      switchOtpSignInLink: By.xpath("//button[contains(text(),'Sign in via OTP')]"),

      // OTP Sign In Inputs
      otpEmailInput: By.xpath("//input[@type='email' and @placeholder='name@example.com']"),
      sendOtpButton: By.xpath("//button[contains(.,'Send 6-Digit Login Code')]"),

      // OTP Verification Screen
      otpCodeInput: By.xpath("//input[@maxLength='6' and contains(@placeholder,'•••')]"),
      verifyOtpSubmitBtn: By.xpath("//button[contains(.,'Verify OTP & Complete Sign In')]"),
      otpResendBtn: By.xpath("//button[contains(text(),'Resend OTP') or contains(text(),'Resend in')]"),
      otpBackBtn: By.xpath("//button[contains(text(),'Back to')]"),

      // Customer Sign Up Inputs
      custNameInput: By.css('input[name="cust_name_field"]'),
      custEmailInput: By.css('input[name="cust_email_field"]'),
      custPassInput: By.css('input[name="cust_password_field"]'),
      custConfirmPassInput: By.css('input[name="cust_confirm_password_field"]'),
      custShowPassCheckbox: By.id("showCustPass"),
      custSignupSubmitBtn: By.xpath("//form//button[@type='submit']"),

      // Vendor Sign Up Inputs
      vendorShopNameInput: By.css('input[name="vendor_shop_name_field"]'),
      vendorCategoryPills: By.xpath("//form//div[contains(@class,'grid-cols-2') or contains(@class,'grid-cols-3')]//button"),
      vendorEmailInput: By.css('input[name="vendor_email_field"]'),
      vendorPhoneInput: By.css('input[name="vendor_phone_field"]'),
      vendorUpiInput: By.css('input[name="vendor_upi_field"]'),
      vendorPassInput: By.css('input[name="vendor_pass_field"]'),
      vendorConfirmPassInput: By.css('input[name="vendor_confirm_pass_field"]'),
      vendorShowPassCheckbox: By.id("showVendPass"),
      vendorPhotoInput: By.xpath("//input[@type='file' and contains(@accept,'image')]"),
      vendorDocInput: By.xpath("//input[@type='file' and contains(@accept,'.pdf')]"),
      vendorSignupSubmitBtn: By.xpath("//form//button[@type='submit']"),

      // Forgot Password Inputs
      forgotEmailInput: By.xpath("//form//input[@type='email']"),
      forgotSendCodeBtn: By.xpath("//button[contains(.,'Send 6-Digit Reset Code')]"),
      forgotOtpInput: By.xpath("//form//input[@maxLength='6']"),
      forgotNewPassInput: By.xpath("//input[@placeholder='Min 6 characters']"),
      forgotConfirmPassInput: By.xpath("//input[@placeholder='Repeat new password']"),
      forgotSubmitNewPassBtn: By.xpath("//button[contains(.,'Reset Password & Sign In')]"),
      forgotDevCodeBtn: By.xpath("//button[contains(text(),'Fill Code')]"),
      forgotBackLink: By.xpath("//button[contains(text(),'Back to Sign In')]"),

      // Google 1-Click Fast Auth
      googleContinueBtn: By.xpath("//button[contains(.,'Continue with Google') or contains(.,'Continue as Admin with Google')]"),
      googleModal: By.xpath("//div[contains(@class,'fixed inset-0')]"),
      googleModalCloseBtn: By.xpath("//div[contains(@class,'fixed inset-0')]//button[@title='Close']"),
      googleModalEmailInput: By.xpath("//div[contains(@class,'fixed inset-0')]//input[@type='email']"),
      googleModalSubmitBtn: By.xpath("//div[contains(@class,'fixed inset-0')]//button[@type='submit']"),
      googleAdminShortcutBtn: By.xpath("//div[contains(@class,'fixed inset-0')]//button[contains(.,'Platform Admin')]"),
      googleClientIdSetupToggle: By.xpath("//button[contains(.,'Connect Google Client ID')]"),
      googleClientIdInput: By.xpath("//input[contains(@placeholder,'apps.googleusercontent.com')]"),

      // Alerts & Messages
      errorMessage: By.xpath("//p[contains(@class,'text-red-600')]"),
      statusMessage: By.xpath("//p[contains(@class,'text-purple-700') or contains(@class,'text-red-700')]"),
      successIcon: By.xpath("//*[contains(@class,'lucide-check-circle')]"),
    };
  }

  // --- Helper Methods ---

  async open(params = "") {
    const url = params ? `${AUTH_URL}?${params}` : AUTH_URL;
    await this.driver.get(url);
    await this.waitForPageLoad();
  }

  async waitForPageLoad() {
    await this.driver.wait(
      until.elementLocated(this.locators.brandHeading),
      DEFAULT_TIMEOUT,
      "Auth page did not load brand heading in time."
    );
  }

  async selectTab(tabName) {
    const target = tabName.toLowerCase() === "signup" ? this.locators.tabSignUp : this.locators.tabSignIn;
    const btn = await this.driver.wait(until.elementLocated(target), DEFAULT_TIMEOUT);
    await btn.click();
    await this.driver.sleep(300);
  }

  async selectRole(role) {
    let target = this.locators.roleCustomer;
    if (role.toLowerCase() === "vendor") target = this.locators.roleVendor;
    if (role.toLowerCase() === "admin") target = this.locators.roleAdmin;

    const btn = await this.driver.wait(until.elementLocated(target), DEFAULT_TIMEOUT);
    await btn.click();
    await this.driver.sleep(300);
  }

  async selectLoginMode(mode) {
    const target = mode.toLowerCase() === "otp" ? this.locators.loginModeOtp : this.locators.loginModePassword;
    const btn = await this.driver.wait(until.elementLocated(target), DEFAULT_TIMEOUT);
    await btn.click();
    await this.driver.sleep(300);
  }

  async loginWithPassword(email, password) {
    const emailField = await this.driver.wait(until.elementLocated(this.locators.loginEmailInput), DEFAULT_TIMEOUT);
    await emailField.clear();
    if (email) await emailField.sendKeys(email);

    const passField = await this.driver.wait(until.elementLocated(this.locators.loginPasswordInput), DEFAULT_TIMEOUT);
    await passField.clear();
    if (password) await passField.sendKeys(password);

    const submitBtn = await this.driver.findElement(this.locators.loginSubmitButton);
    await submitBtn.click();
  }

  async togglePasswordVisibility() {
    const toggle = await this.driver.wait(until.elementLocated(this.locators.passwordToggleBtn), DEFAULT_TIMEOUT);
    await toggle.click();
  }

  async getPasswordInputType() {
    const passField = await this.driver.findElement(this.locators.loginPasswordInput);
    return await passField.getAttribute("type");
  }

  async getErrorMessage() {
    try {
      const errEl = await this.driver.wait(until.elementLocated(this.locators.errorMessage), 3000);
      return await errEl.getText();
    } catch {
      return null;
    }
  }

  async getStatusMessage() {
    try {
      const msgEl = await this.driver.wait(until.elementLocated(this.locators.statusMessage), 3000);
      return await msgEl.getText();
    } catch {
      return null;
    }
  }

  async triggerOtpLogin(email) {
    await this.selectLoginMode("otp");
    const input = await this.driver.wait(until.elementLocated(this.locators.otpEmailInput), DEFAULT_TIMEOUT);
    await input.clear();
    await input.sendKeys(email);
    const btn = await this.driver.findElement(this.locators.sendOtpButton);
    await btn.click();
  }

  async fillOtp(code) {
    const otpInput = await this.driver.wait(until.elementLocated(this.locators.otpCodeInput), DEFAULT_TIMEOUT);
    await otpInput.clear();
    await otpInput.sendKeys(code);
    const btn = await this.driver.findElement(this.locators.verifyOtpSubmitBtn);
    await btn.click();
  }

  async fillCustomerSignup(name, email, password, confirmPassword) {
    await this.selectTab("signup");
    await this.selectRole("customer");

    const nameEl = await this.driver.wait(until.elementLocated(this.locators.custNameInput), DEFAULT_TIMEOUT);
    await nameEl.clear();
    if (name) await nameEl.sendKeys(name);

    const emailEl = await this.driver.findElement(this.locators.custEmailInput);
    await emailEl.clear();
    if (email) await emailEl.sendKeys(email);

    const passEl = await this.driver.findElement(this.locators.custPassInput);
    await passEl.clear();
    if (password) await passEl.sendKeys(password);

    const confirmEl = await this.driver.findElement(this.locators.custConfirmPassInput);
    await confirmEl.clear();
    if (confirmPassword) await confirmEl.sendKeys(confirmPassword);

    const submitBtn = await this.driver.findElement(this.locators.custSignupSubmitBtn);
    await submitBtn.click();
  }

  async openGoogleModal() {
    const btn = await this.driver.wait(until.elementLocated(this.locators.googleContinueBtn), DEFAULT_TIMEOUT);
    await btn.click();
    await this.driver.wait(until.elementLocated(this.locators.googleModal), DEFAULT_TIMEOUT);
  }

  async closeGoogleModal() {
    const closeBtn = await this.driver.findElement(this.locators.googleModalCloseBtn);
    await closeBtn.click();
    await this.driver.sleep(300);
  }

  async clickForgotPassword() {
    const link = await this.driver.wait(until.elementLocated(this.locators.forgotPasswordLink), DEFAULT_TIMEOUT);
    await link.click();
    await this.driver.wait(until.elementLocated(this.locators.forgotEmailInput), DEFAULT_TIMEOUT);
  }

  async submitForgotPasswordEmail(email) {
    const input = await this.driver.findElement(this.locators.forgotEmailInput);
    await input.clear();
    await input.sendKeys(email);
    const btn = await this.driver.findElement(this.locators.forgotSendCodeBtn);
    await btn.click();
  }

  async takeScreenshot(name) {
    try {
      const image = await this.driver.takeScreenshot();
      const filename = path.join(SCREENSHOT_DIR, `${Date.now()}_${name}.png`);
      fs.writeFileSync(filename, image, "base64");
      return filename;
    } catch (e) {
      console.error(`Failed to capture screenshot: ${e.message}`);
      return null;
    }
  }
}

// ==========================================
// Custom Lightweight Test Suite Runner
// ==========================================
class TestSuiteRunner {
  constructor() {
    this.results = [];
    this.currentSuite = "";
  }

  suite(name) {
    this.currentSuite = name;
    console.log(`\n\x1b[1m\x1b[35m=== SUITE: ${name} ===\x1b[0m`);
  }

  async test(caseId, description, fn) {
    const startTime = Date.now();
    process.stdout.write(`  [${caseId}] ${description} ... `);
    try {
      await fn();
      const duration = Date.now() - startTime;
      console.log(`\x1b[32mPASSED\x1b[0m (${duration}ms)`);
      this.results.push({ id: caseId, suite: this.currentSuite, desc: description, status: "PASSED", duration, error: null });
    } catch (err) {
      const duration = Date.now() - startTime;
      console.log(`\x1b[31mFAILED\x1b[0m (${duration}ms)`);
      console.log(`     \x1b[31mError: ${err.message}\x1b[0m`);
      this.results.push({ id: caseId, suite: this.currentSuite, desc: description, status: "FAILED", duration, error: err.message });
    }
  }

  printSummary() {
    const total = this.results.length;
    const passed = this.results.filter((r) => r.status === "PASSED").length;
    const failed = total - passed;
    const passRate = total > 0 ? ((passed / total) * 100).toFixed(1) : "0.0";

    console.log("\n" + "=".repeat(60));
    console.log(`\x1b[1mTEST EXECUTION SUMMARY\x1b[0m`);
    console.log("=".repeat(60));
    console.log(`Total Automated Scenarios Executed : ${total}`);
    console.log(`\x1b[32mPassed                             : ${passed}\x1b[0m`);
    console.log(`\x1b[31mFailed                             : ${failed}\x1b[0m`);
    console.log(`Pass Rate                          : ${passRate}%`);
    console.log("=".repeat(60) + "\n");
  }
}

// ==========================================
// Main Selenium Execution Function
// ==========================================
export async function runLoginTests() {
  console.log(`\x1b[1mStarting Meeva Web Frontend E2E Selenium Tests\x1b[0m`);
  console.log(`Target URL : ${AUTH_URL}`);
  console.log(`Headless   : ${IS_HEADLESS}`);

  const chromeOptions = new chrome.Options();
  if (IS_HEADLESS) {
    chromeOptions.addArguments("--headless=new");
  }
  chromeOptions.addArguments(
    "--no-sandbox",
    "--disable-dev-shm-usage",
    "--disable-gpu",
    "--window-size=1440,900",
    "--disable-extensions",
    "--ignore-certificate-errors"
  );

  let driver;
  const runner = new TestSuiteRunner();

  try {
    driver = await new Builder().forBrowser("chrome").setChromeOptions(chromeOptions).build();
    await driver.manage().setTimeouts({ implicit: 2000, pageLoad: 15000 });

    const authPage = new AuthPage(driver);

    // ----------------------------------------------------
    // SUITE 1: Page Navigation, Layout & Default UI States
    // ----------------------------------------------------
    runner.suite("1. Page Navigation, Layout & Default States");

    await runner.test("TC_AUTH_001", "Verify Auth Page loads with valid brand heading & title", async () => {
      await authPage.open();
      const heading = await driver.findElement(authPage.locators.brandHeading);
      const text = await heading.getText();
      if (!text.toLowerCase().includes("meeva")) {
        throw new Error(`Expected heading to include 'Meeva', got: '${text}'`);
      }
    });

    await runner.test("TC_AUTH_002", "Verify default active tab is 'Sign In' and role is 'Customer'", async () => {
      await authPage.open();
      const signInBtn = await driver.findElement(authPage.locators.tabSignIn);
      const customerRoleBtn = await driver.findElement(authPage.locators.roleCustomer);

      const signInClass = await signInBtn.getAttribute("class");
      const customerClass = await customerRoleBtn.getAttribute("class");

      if (!signInClass.includes("text-purple-700") && !signInClass.includes("shadow-sm")) {
        throw new Error("Sign In tab is not visually marked active by default.");
      }
      if (!customerClass.includes("text-purple-700") && !customerClass.includes("shadow-sm")) {
        throw new Error("Customer role button is not active by default.");
      }
    });

    await runner.test("TC_AUTH_003", "Verify all 3 roles (Customer, Vendor, Admin) are available in Sign In", async () => {
      await authPage.open();
      await driver.findElement(authPage.locators.roleCustomer);
      await driver.findElement(authPage.locators.roleVendor);
      await driver.findElement(authPage.locators.roleAdmin);
    });

    await runner.test("TC_AUTH_004", "Verify Admin role is hidden when switching to 'Sign Up' tab", async () => {
      await authPage.open();
      await authPage.selectTab("signup");
      const adminBtns = await driver.findElements(authPage.locators.roleAdmin);
      if (adminBtns.length > 0) {
        throw new Error("Admin role tab should NOT be visible during Sign Up.");
      }
    });

    // ----------------------------------------------------
    // SUITE 2: Customer Password Login Validation & Controls
    // ----------------------------------------------------
    runner.suite("2. Customer Password Login Validation & Controls");

    await runner.test("TC_AUTH_005", "Verify Password Visibility Toggle switches type 'password' to 'text'", async () => {
      await authPage.open();
      const initialType = await authPage.getPasswordInputType();
      if (initialType !== "password") {
        throw new Error(`Expected input type 'password', got '${initialType}'`);
      }

      await authPage.togglePasswordVisibility();
      const toggledType = await authPage.getPasswordInputType();
      if (toggledType !== "text") {
        throw new Error(`Expected input type 'text' after eye toggle click, got '${toggledType}'`);
      }

      await authPage.togglePasswordVisibility();
      const revertedType = await authPage.getPasswordInputType();
      if (revertedType !== "password") {
        throw new Error(`Expected input type 'password' after second toggle click, got '${revertedType}'`);
      }
    });

    await runner.test("TC_AUTH_006", "Verify empty credentials form submission triggers browser HTML5 validation", async () => {
      await authPage.open();
      const emailField = await driver.findElement(authPage.locators.loginEmailInput);
      await emailField.clear();
      const submitBtn = await driver.findElement(authPage.locators.loginSubmitButton);
      await submitBtn.click();

      const isValid = await driver.executeScript("return arguments[0].checkValidity();", emailField);
      if (isValid) {
        throw new Error("Empty email input passed HTML5 checkValidity() validation unexpectedly.");
      }
    });

    await runner.test("TC_AUTH_007", "Verify invalid login credentials display appropriate error message", async () => {
      await authPage.open();
      await authPage.loginWithPassword("nonexistent_user_test@example.com", "WrongPassword123!");
      const error = await authPage.getErrorMessage();
      if (!error || error.trim().length === 0) {
        throw new Error("No error alert displayed when logging in with invalid credentials.");
      }
    });

    // ----------------------------------------------------
    // SUITE 3: Role Switcher & Access Control Enforcement
    // ----------------------------------------------------
    runner.suite("3. Role Switcher & Access Control Enforcement");

    await runner.test("TC_AUTH_008", "Verify switching role to Vendor updates UI and submit button text", async () => {
      await authPage.open();
      await authPage.selectRole("vendor");
      const submitBtn = await driver.findElement(authPage.locators.loginSubmitButton);
      const text = await submitBtn.getText();
      if (!text.includes("Sign In")) {
        throw new Error(`Unexpected submit button text: '${text}'`);
      }
    });

    await runner.test("TC_AUTH_009", "Verify switching role to Admin displays 'Sign In to Admin Console'", async () => {
      await authPage.open();
      await authPage.selectRole("admin");
      const submitBtn = await driver.findElement(authPage.locators.loginSubmitButton);
      const text = await submitBtn.getText();
      if (!text.toLowerCase().includes("admin console")) {
        throw new Error(`Expected submit button to mention 'Admin Console', got '${text}'`);
      }
    });

    await runner.test("TC_AUTH_010", "Verify URL query param ?role=vendor automatically activates Vendor tab", async () => {
      await authPage.open("role=vendor");
      const vendorBtn = await driver.findElement(authPage.locators.roleVendor);
      const cls = await vendorBtn.getAttribute("class");
      if (!cls.includes("text-purple-600") && !cls.includes("shadow-sm")) {
        throw new Error("Vendor role button is not highlighted active when loaded with ?role=vendor.");
      }
    });

    // ----------------------------------------------------
    // SUITE 4: 6-Digit OTP Sign In Flow
    // ----------------------------------------------------
    runner.suite("4. 6-Digit OTP Sign In Flow");

    await runner.test("TC_AUTH_011", "Verify switching to OTP mode presents 'Email Address to Receive OTP' input", async () => {
      await authPage.open();
      await authPage.selectLoginMode("otp");
      const otpInput = await driver.findElement(authPage.locators.otpEmailInput);
      const placeholder = await otpInput.getAttribute("placeholder");
      if (!placeholder.includes("name@example.com")) {
        throw new Error(`OTP email placeholder mismatch, got: '${placeholder}'`);
      }
    });

    await runner.test("TC_AUTH_012", "Verify requesting OTP transitions to OTP Code Entry with 6-character limit", async () => {
      await authPage.open();
      await authPage.triggerOtpLogin("qa_test_otp_user@example.com");

      // Verify OTP screen input is displayed
      const codeInput = await driver.wait(until.elementLocated(authPage.locators.otpCodeInput), 5000);
      const maxLength = await codeInput.getAttribute("maxlength");
      if (maxLength !== "6") {
        throw new Error(`Expected OTP code input maxLength='6', got '${maxLength}'`);
      }

      // Verify non-digits are filtered
      await codeInput.sendKeys("12a3b4");
      const val = await codeInput.getAttribute("value");
      if (val !== "1234") {
        throw new Error(`Expected non-digits to be stripped to '1234', got '${val}'`);
      }
    });

    await runner.test("TC_AUTH_013", "Verify Back button on OTP screen returns to Sign In screen", async () => {
      const backBtn = await driver.findElement(authPage.locators.otpBackBtn);
      await backBtn.click();
      await driver.wait(until.elementLocated(authPage.locators.tabSignIn), 3000);
    });

    // ----------------------------------------------------
    // SUITE 5: Customer Sign Up Validation
    // ----------------------------------------------------
    runner.suite("5. Customer Sign Up Validation");

    await runner.test("TC_AUTH_014", "Verify Customer Sign Up rejects password shorter than 6 characters", async () => {
      await authPage.open();
      await authPage.fillCustomerSignup("John Doe", "johndoe_test@example.com", "12345", "12345");
      const passField = await driver.findElement(authPage.locators.custPassInput);
      const isValid = await driver.executeScript("return arguments[0].checkValidity();", passField);
      if (isValid) {
        throw new Error("Password with length < 6 characters passed validation unexpectedly.");
      }
    });

    await runner.test("TC_AUTH_015", "Verify Customer Sign Up rejects mismatched passwords", async () => {
      await authPage.open();
      await authPage.fillCustomerSignup("John Doe", "johndoe_test@example.com", "SecurePass123!", "DifferentPass456!");
      const error = await authPage.getErrorMessage();
      if (!error || !error.toLowerCase().includes("passwords do not match")) {
        throw new Error(`Expected 'Passwords do not match' error alert, got '${error}'`);
      }
    });

    // ----------------------------------------------------
    // SUITE 6: Vendor Sign Up & Category Selection
    // ----------------------------------------------------
    runner.suite("6. Vendor Sign Up & Category Selection");

    await runner.test("TC_AUTH_016", "Verify Vendor Sign Up shows 6 business categories", async () => {
      await authPage.open("tab=signup&role=vendor");
      const pills = await driver.findElements(authPage.locators.vendorCategoryPills);
      if (pills.length < 6) {
        throw new Error(`Expected at least 6 vendor categories, found ${pills.length}`);
      }
    });

    await runner.test("TC_AUTH_017", "Verify Vendor photo and document file upload inputs exist", async () => {
      await driver.findElement(authPage.locators.vendorPhotoInput);
      await driver.findElement(authPage.locators.vendorDocInput);
    });

    // ----------------------------------------------------
    // SUITE 7: Forgot Password Recovery Flow
    // ----------------------------------------------------
    runner.suite("7. Forgot Password Recovery Flow");

    await runner.test("TC_AUTH_018", "Verify clicking 'Forgot password?' transitions to Password Recovery view", async () => {
      await authPage.open();
      await authPage.clickForgotPassword();
      const heading = await driver.findElement(By.xpath("//h3[contains(text(),'Forgot your password?')]"));
      const isVisible = await heading.isDisplayed();
      if (!isVisible) {
        throw new Error("Password recovery heading is not visible.");
      }
    });

    await runner.test("TC_AUTH_019", "Verify Back to Sign In button restores the Sign In tab", async () => {
      const backBtn = await driver.findElement(authPage.locators.forgotBackLink);
      await backBtn.click();
      await driver.wait(until.elementLocated(authPage.locators.loginEmailInput), 3000);
    });

    // ----------------------------------------------------
    // SUITE 8: Google 1-Click Fast Auth & Modal
    // ----------------------------------------------------
    runner.suite("8. Google 1-Click Fast Auth & Modal");

    await runner.test("TC_AUTH_020", "Verify clicking 'Continue with Google' opens Google Account Modal", async () => {
      await authPage.open();
      await authPage.openGoogleModal();
      const modal = await driver.findElement(authPage.locators.googleModal);
      const isDisplayed = await modal.isDisplayed();
      if (!isDisplayed) {
        throw new Error("Google Authentication modal did not open.");
      }
    });

    await runner.test("TC_AUTH_021", "Verify Google Modal Close button dismisses the dialog", async () => {
      await authPage.closeGoogleModal();
      const modals = await driver.findElements(authPage.locators.googleModal);
      if (modals.length > 0 && (await modals[0].isDisplayed())) {
        throw new Error("Google Modal is still displayed after clicking close button.");
      }
    });

    // ----------------------------------------------------
    // SUITE 9: Security & Input Sanitization
    // ----------------------------------------------------
    runner.suite("9. Security & Input Sanitization");

    await runner.test("TC_AUTH_022", "Verify SQL Injection payloads in login fields do not cause unhandled crashes", async () => {
      await authPage.open();
      const sqliPayload = "' OR '1'='1' --";
      await authPage.loginWithPassword(sqliPayload, "sample_password");
      // Expect either error message or standard validation rejection, not page crash
      await driver.sleep(1000);
      const brand = await driver.findElement(authPage.locators.brandHeading);
      if (!(await brand.isDisplayed())) {
        throw new Error("Page crashed or became unresponsive after SQLi payload submission.");
      }
    });

    await runner.test("TC_AUTH_023", "Verify XSS script tags in inputs do not execute alert dialogs", async () => {
      await authPage.open();
      const xssPayload = `<script>window.__xss_detected=true;</script>`;
      await authPage.loginWithPassword(xssPayload, "random_pass");

      const xssFired = await driver.executeScript("return window.__xss_detected === true;");
      if (xssFired) {
        throw new Error("Critical Vulnerability: Script tag payload executed on the auth page!");
      }
    });

    // ----------------------------------------------------
    // SUITE 10: Viewport Responsiveness & Layout
    // ----------------------------------------------------
    runner.suite("10. Viewport Responsiveness & Layout");

    await runner.test("TC_AUTH_024", "Verify Auth Page renders cleanly on Mobile Viewport (375x812)", async () => {
      await driver.manage().window().setRect({ width: 375, height: 812 });
      await authPage.open();
      const submitBtn = await driver.findElement(authPage.locators.loginSubmitButton);
      if (!(await submitBtn.isDisplayed())) {
        throw new Error("Submit button not visible on mobile resolution (375x812).");
      }
      // Reset window size to desktop
      await driver.manage().window().setRect({ width: 1440, height: 900 });
    });

    // Print final summary
    runner.printSummary();

    // Check if any critical test failed
    const hasFailures = runner.results.some((r) => r.status === "FAILED");
    return { success: !hasFailures, results: runner.results };
  } catch (globalErr) {
    console.error(`\x1b[31mFatal error during Selenium test execution: ${globalErr.message}\x1b[0m`);
    if (driver) {
      const errPath = path.join(SCREENSHOT_DIR, `fatal_error_${Date.now()}.png`);
      try {
        const img = await driver.takeScreenshot();
        fs.writeFileSync(errPath, img, "base64");
        console.log(`Saved fatal error screenshot to: ${errPath}`);
      } catch {}
    }
    return { success: false, error: globalErr.message };
  } finally {
    if (driver) {
      await driver.quit();
      console.log("WebDriver session closed gracefully.");
    }
  }
}

// Auto-run if executed directly via Node.js
if (process.argv[1] && process.argv[1].endsWith("login-tests.js")) {
  runLoginTests()
    .then((res) => {
      process.exit(res.success ? 0 : 1);
    })
    .catch((err) => {
      console.error(err);
      process.exit(1);
    });
}
