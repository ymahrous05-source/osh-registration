/**
 * ============================================================================
 *  Osar Sonaa Al-Hayah — Registration backend (Google Apps Script)
 * ============================================================================
 *  This is the server-side code behind:
 *    - dys_form.html      (public registration form, sends doPost)
 *    - dys_dashboard.html (admin dashboard, calls doGet?action=list, plus the
 *                          new "الإعدادات" tab which manages the registration
 *                          window, form title, sheet name, and the logo image)
 *
 *  ⚠️ READ THIS BEFORE DEPLOYING ⚠️
 *  If you already have real registrations in a Google Sheet:
 *    1. Open THAT sheet first → Extensions ▸ Apps Script. If code shows up
 *       there, your old backend isn't actually lost — it's attached to the
 *       sheet. Don't replace it blindly; compare it with this file instead.
 *    2. Make a copy of the sheet (File ▸ Make a copy) before testing this,
 *       so you can experiment without risking real data.
 *    3. Make sure SHEET_NAME below matches your real tab name, and that
 *       row 1 of your sheet has headers matching HEADERS below — in the
 *       same order. Run checkHeaders() (see bottom of this file) from the
 *       Apps Script editor to get a report instead of guessing.
 *
 *  ✅ UPGRADING FROM THE OLDER VERSION OF THIS FILE
 *  This version is backward compatible with your existing data: on first
 *  run it keeps using whatever sheet tab is named SHEET_NAME below (your
 *  real registrations stay exactly where they are). The new "cycle" system
 *  (registration windows + auto-created sheets) only kicks in once you set
 *  it up from the dashboard's Settings tab — nothing changes until you do.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// CONFIG — the only section you should normally need to touch
// ---------------------------------------------------------------------------

// Bump this string any time you paste in a new version of this file. After
// deploying, open ?action=diagnostics&password=... (or the "🩺 تشخيص" button
// in the dashboard) — if codeVersion doesn't match what you expect, the web
// app is still running an OLD deployment and you need Deploy ▸ Manage
// deployments ▸ Edit (✏️) ▸ New version ▸ Deploy (see BACKEND_SETUP_STEPS.md
// step 7). This single check rules out the #1 cause of "I edited the code
// but nothing changed."
const CODE_VERSION = "2026-08-25-checkin-scanner";

// ---------------------------------------------------------------------------
// Google Sign-In (optional — the old shared password keeps working forever
// alongside this, nothing breaks if you never set this up)
// ---------------------------------------------------------------------------
// Paste your OAuth Client ID here (from Google Cloud Console ▸ APIs &
// Services ▸ Credentials — see BACKEND_SETUP_STEPS.md). Leave the
// placeholder as-is and the "تسجيل دخول بجوجل" button just won't work yet —
// everything else (password login) is unaffected.
const GOOGLE_CLIENT_ID = ""; // TODO: paste your own OAuth Client ID here (Google Cloud Console) if you want "Sign in with Google" — see comment above. Password login works fine without it.

// How long a Google-signed-in session stays valid before that person has to
// sign in again (they'll just see the "Sign in with Google" button reappear
// — nothing is lost, no data at risk).
const SESSION_TTL_MS = 6 * 60 * 60 * 1000; // 6 hours

// ---------------------------------------------------------------------------
// Outgoing email sender identity
// ---------------------------------------------------------------------------
// All emails (confirmations, certificates, new-access notifications) go out
// "from" this address instead of whichever Google account happens to run
// the script.
//
// ⚠️ REQUIRES a one-time Gmail setup step first, or every send will fail:
// the Google account that owns/runs this Apps Script project must add
// FROM_EMAIL as a verified "Send mail as" address —
//   Gmail ⚙️ (that account's inbox) ▸ See all settings ▸ Accounts and Import
//   ▸ "Send mail as" ▸ Add another email address ▸ enter FROM_EMAIL ▸ Gmail
//   sends a verification link/code to FROM_EMAIL's own inbox ▸ confirm it.
// Once verified, this takes effect immediately — no redeploy needed beyond
// the normal "paste the value, save, New deployment" flow.
//
// Leave FROM_EMAIL as "" to skip all this and just send from the script
// owner's own address, exactly like before this feature existed.
const FROM_EMAIL = ""; // TODO: set a verified "send mail as" address for أسرة صناع الحياة, or leave "" to send from the script owner's own Google account
const FROM_NAME = "أسرة صناع الحياة";

// Central email sender — every outgoing email in this file goes through
// here, so the "from" identity only ever needs to be changed in ONE place.
// `extraOptions` can add things like {attachments: [...]} on top.
function sendEmail_(to, subject, body, extraOptions) {
  const options = Object.assign({ name: FROM_NAME }, extraOptions || {});
  if (FROM_EMAIL) options.from = FROM_EMAIL;
  GmailApp.sendEmail(to, subject, body, options);
}

// ---------------------------------------------------------------------------
// AI chat assistant (optional — the FAQ quick-buttons in both chat widgets
// work with zero setup; only the free-text "اكتب سؤالك" box needs a key)
// ---------------------------------------------------------------------------
// Gemini is tried FIRST (free tier — get a key in 2 minutes at
// aistudio.google.com ▸ Get API key). Grok is the FALLBACK if Gemini fails
// or isn't configured (x.ai — needs billing enabled, no generous free tier,
// so it's optional). Leave either/both placeholders as-is to skip that
// provider entirely — the chat still works with just one, or falls back to
// "use the quick questions instead" if neither is set.
const GEMINI_API_KEY = "AQ.Ab8RN6JHY6laQxjm-W-_q3-PthfhbDrTfvY9dFth6cOJiwUD8g";
const GROK_API_KEY = "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";

// Simple daily message cap (shared across all forms) so a misconfigured
// widget or a bored visitor can't run up an unexpected API bill overnight.
// Resets automatically at midnight (script timezone) — raise this if 300/day
// is genuinely too low for your traffic.
const AI_DAILY_LIMIT = 300;

// The four independently-grantable admin capabilities. "viewData" is always
// true for every account (there's no point in an account that can log in
// and see nothing) — the other four are what actually differ per account.
// Add a new capability here + wire it into ACTION_PERMISSIONS below and the
// dashboard's account-editor checkboxes if you ever need a 6th one.
const PERMISSION_KEYS = ["manageSettings", "manageFields", "manageCertificates", "manageAccounts"];

const SHEET_NAME = "Registrations"; // <-- change to match your actual tab name (used as the default/first cycle's sheet)

// The Sheet this backend reads/writes. SpreadsheetApp.getActiveSpreadsheet()
// only works when a script is BOUND to a Sheet and running from inside its
// UI — it returns null for a deployed Web App (which is how this backend
// actually runs), causing "Cannot read properties of null" errors. Using
// openById() with an explicit ID works from any context, bound or not.
// Get this ID from your Sheet's URL: .../spreadsheets/d/THIS_PART/edit
const SHEET_ID = "1H3fl2szORZhW5JmT5R2OY-rX0RAxuysmmnBDhH85ZE0";
function getSpreadsheet_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// Column order written to the sheet. Keep this order in sync with HEADERS —
// index i of HEADERS must correspond to index i of the row array built in
// buildRow_(). The dashboard finds columns by header NAME (not position), so
// reordering here is safe as long as HEADERS and buildRow_() stay matched.
// ---------------------------------------------------------------------------
// Extended field registry — the "literally every field a form might need"
// set. Each of these behaves exactly like the original 9 toggleable fields
// (age, gender, phone, ...) below: shown/hidden and required/optional from
// the "🧩 حقول الاستمارة" settings card, with no code changes needed. They're
// kept in a SEPARATE object from TOGGLEABLE_FIELDS (rather than merged in)
// so the original 9 fields' hand-written validation further down stays
// completely untouched — these newer ones are validated generically instead
// (see the EXTRA_FIELDS loop inside validatePayload_), based on `type`:
//   "text"     — any non-empty string passes (min-length 1)
//   "textarea" — same as text, just a bigger box on the form
//   "select"   — value must be one of `options`
//   "date"     — must look like YYYY-MM-DD
//   "checkbox" — stored/validated as "Yes"/"No", same convention as the
//                existing graduate/hasJob fields
// `section` groups fields on both the dashboard's config card and the
// public form — a section with zero enabled fields (all disabled here, and
// no custom fields assigned to it) simply never renders, per your request
// that empty pages/sections shouldn't show up in the form at all.
const FIELD_SECTIONS = {
  personal: "بيانات شخصية إضافية",
  education: "بيانات تعليمية ووظيفية إضافية",
  contact: "بيانات تواصل إضافية",
  entity: "بيانات داخل الكيان إضافية",
  other: "حقول تانية",
  custom: "حقول مخصصة",
};

const EXTRA_FIELDS = {
  address:        { section: "personal", label: "العنوان بالتفصيل", type: "text", defaultRequired: false },
  birthDate:      { section: "personal", label: "تاريخ الميلاد", type: "date", defaultRequired: false },
  maritalStatus:  { section: "personal", label: "الحالة الاجتماعية", type: "select", options: ["أعزب", "متزوج", "مطلق", "أرمل"], defaultRequired: false },
  governorate:    { section: "personal", label: "المحافظة", type: "text", defaultRequired: false },
  academicYear:   { section: "education", label: "الفرقة الدراسية", type: "text", defaultRequired: false },
  gradeLevel:     { section: "education", label: "التقدير الدراسي", type: "text", defaultRequired: false },
  facebook:       { section: "contact", label: "رابط الفيسبوك", type: "text", defaultRequired: false },
  instagram:      { section: "contact", label: "يوزر الانستجرام", type: "text", defaultRequired: false },
  emergencyName:  { section: "contact", label: "اسم شخص للطوارئ", type: "text", defaultRequired: false },
  emergencyPhone: { section: "contact", label: "رقم شخص للطوارئ", type: "text", defaultRequired: false },
  howHeard:       { section: "entity", label: "عرفت عننا إزاي؟", type: "text", defaultRequired: false },
  prevVolunteer:  { section: "entity", label: "خبرة تطوعية سابقة", type: "textarea", defaultRequired: false },
  motivation:     { section: "entity", label: "ليه عايز تنضم؟", type: "textarea", defaultRequired: false },
  skills:         { section: "entity", label: "مهارات أو اهتمامات", type: "text", defaultRequired: false },
  availability:   { section: "entity", label: "الأوقات المتاحة للتطوع", type: "text", defaultRequired: false },
  tshirtSize:     { section: "other", label: "مقاس التيشيرت", type: "select", options: ["S", "M", "L", "XL", "XXL"], defaultRequired: false },
  notes:          { section: "other", label: "ملاحظات إضافية", type: "textarea", defaultRequired: false },
};

// address reuses the pre-existing "Address" column (it was always in
// HEADERS, just unconditionally blank — see buildRow_) so it's NOT in
// EXTRA_HEADERS below; everything else here is a genuinely new column.
const EXTRA_HEADERS = [
  "Birth Date", "Marital Status", "Governorate", "Academic Year", "Grade Level",
  "Facebook", "Instagram", "Emergency Contact Name", "Emergency Contact Phone",
  "How Heard", "Previous Volunteering", "Motivation", "Skills", "Availability",
  "Tshirt Size", "Notes",
];
const EXTRA_HEADER_KEYS = Object.keys(EXTRA_FIELDS).filter(k => k !== "address");
const HEADER_TO_EXTRA_FIELD = {};
EXTRA_HEADERS.forEach((h, i) => { HEADER_TO_EXTRA_FIELD[h] = EXTRA_HEADER_KEYS[i]; });

const CUSTOM_FIELDS_HEADER = "Custom Fields (JSON)";

const HEADERS = [
  "Timestamp",
  "Membership No",
  "Name",
  "Age",
  "Gender",
  "National ID",
  "Phone",
  "Whatsapp",
  "Email",
  "Address",
  "Faculty",
  "Graduate",
  "Role in Entity",   // ← this is the "committee" field from the form
  "Has Job",
  "Current Job",
].concat(EXTRA_HEADERS).concat([CUSTOM_FIELDS_HEADER]).concat(["Photo URL", "Video URL", "Checked In At"]);

const MIN_FILL_MS = 2500; // mirrors the frontend's own MIN_FILL_MS anti-bot check
const MEMBERSHIP_PREFIX = "OSH";

// ---------------------------------------------------------------------------
// Configurable fields — controlled from the dashboard's "🧩 حقول الاستمارة"
// settings card. Each of these can be shown/hidden on the form, and marked
// required or optional, WITHOUT touching any code or the sheet's columns
// (the column stays in HEADERS either way — it's just left blank when a
// field is disabled or skipped). "Name" and "National ID" are intentionally
// NOT in this list: the whole duplicate-check + membership system depends
// on them, so they always stay shown and required.
const TOGGLEABLE_FIELDS = {
  age:       { section: "personal", label: "العمر", defaultRequired: true },
  gender:    { section: "personal", label: "النوع", defaultRequired: true },
  phone:     { section: "contact", label: "رقم الهاتف", defaultRequired: true },
  whatsapp:  { section: "contact", label: "رقم الواتساب", defaultRequired: true },
  email:     { section: "contact", label: "البريد الإلكتروني", defaultRequired: false },
  faculty:   { section: "education", label: "الكلية / المدرسة", defaultRequired: true },
  graduate:  { section: "education", label: "هل أنت خريج؟", defaultRequired: true },
  committee: { section: "entity", label: "صفتك داخل الكيان", defaultRequired: true },
  hasJob:    { section: "education", label: "هل تعمل حاليًا؟", defaultRequired: true },
};

// The personal-photo field is file-typed (not text/select/etc. like
// EXTRA_FIELDS), so it's validated and uploaded through its own dedicated
// path (see handleSubmit_'s photo-upload step + validatePayload_) rather
// than the generic EXTRA_FIELDS loop — but it still shows up in
// ALL_BUILTIN_FIELDS so the dashboard can enable/require it exactly like
// every other field.
const PHOTO_FIELD = { section: "personal", label: "صورة شخصية", type: "photo", defaultRequired: false };

// Same idea as PHOTO_FIELD, but for an optional short video — uploaded
// straight from the browser to Cloudinary (never through Apps Script, which
// can't handle large files well), with only the resulting URL sent to the
// backend and written to its own "Video URL" column. See dys_form.html's
// renderDynamicField_ ("video" branch) for the upload logic, and
// BACKEND_SETUP_STEPS.md's "Cloudinary" section for the one-time setup.
const VIDEO_FIELD = { section: "personal", label: "فيديو (اختياري)", type: "video", defaultRequired: false };

// Every built-in field (original 9 + the newer 17 + the photo field) — this
// is what the dashboard's "🧩 حقول الاستمارة" card and
// getFieldConfig_/handleSaveFieldConfig_ iterate over. The original 9 keep
// their bespoke hand-written validation in validatePayload_ (unchanged);
// EXTRA_FIELDS are validated generically.
const ALL_BUILTIN_FIELDS = Object.assign({}, TOGGLEABLE_FIELDS, EXTRA_FIELDS, { photo: PHOTO_FIELD, video: VIDEO_FIELD });


// ---------------------------------------------------------------------------
// ENTRY POINTS
// ---------------------------------------------------------------------------

function doGet(e) {
  const action = (e.parameter.action || "").trim();

  try {
    if (action === "list") return handleList_(e);
    if (action === "checkNid") return handleCheckNid_(e);
    if (action === "publicConfig") return handlePublicConfig_(e);
    if (action === "getConfig") return handleGetConfig_(e);
    if (action === "listCycles") return handleListCycles_(e);
    if (action === "diagnostics") return handleDiagnostics_(e);
    if (action === "verifyMember") return handleVerifyMember_(e); // event check-in scanner — read-only lookup
    if (action === "listForms") return handleListForms_(e);
    if (action === "approveAccess") return handleReviewAccess_(e, true);
    if (action === "rejectAccess") return handleReviewAccess_(e, false);

    return jsonOutput_({
      status: "success",
      message: "Osar Sonaa Al-Hayah backend is running. Use ?action=list, ?action=checkNid, or ?action=publicConfig.",
    });
  } catch (err) {
    return jsonOutput_({ status: "error", message: String(err) });
  }
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      return jsonOutput_({ status: "error", message: "No data received" });
    }

    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (err) {
      return jsonOutput_({ status: "error", message: "Invalid JSON" });
    }

    // Event check-in scanner — separate from ADMIN_ACTIONS below because it
    // only needs "viewData" (any valid password), not a specific manage*
    // permission; see handleCheckin_ for its own auth check.
    if (payload && payload.action === "checkin") return handleCheckin_(payload);

    // Google Sign-In — this IS the login step itself, so (unlike everything
    // else below) it can't require a password/session first; see
    // handleGoogleLogin_ for its own verification of the Google ID token.
    if (payload && payload.action === "googleLogin") return handleGoogleLogin_(payload);

    // AI chat widget (form + dashboard). Its own auth check is inside
    // handleChatAI_ (only the dashboard context requires a password/session
    // — the form-facing widget is public, same as the registration form
    // itself, and never sees any registrant data — see buildFormAiPrompt_).
    if (payload && payload.action === "chatAI") return handleChatAI_(payload);

    // Admin-only actions (settings panel in the dashboard) — every one of
    // these re-checks the password itself, so nothing here is trusted blindly.
    const ADMIN_ACTIONS = [
      "saveConfig", "uploadLogo", "removeLogo",
      "uploadCertTemplate", "removeCertTemplate",
      "sendCertificate", "sendCertificatesBulk", "sendTestCertificate",
      "saveFieldConfig",
      "listAdminAccounts", "addAdminAccount", "removeAdminAccount", "reviewAccess", "updateAccountPermissions",
      "exportExcel", "getActivityLog",
      "addForm", "renameForm", "archiveForm",
    ];
    if (payload && ADMIN_ACTIONS.indexOf(payload.action) > -1) {
      return handleAdminAction_(payload);
    }

    // Anything else is treated as a normal registration submission.
    return handleSubmit_(payload || {});
  } catch (err) {
    return jsonOutput_({ status: "error", message: String(err) });
  }
}


// ---------------------------------------------------------------------------
// action=list  (admin dashboard — table + charts)
// ---------------------------------------------------------------------------

function handleList_(e) {
  const password = e.parameter.password || "";
  if (!getAdminAccounts_().length) {
    return jsonOutput_({ status: "error", message: "لسه محددتش كلمة سر الأدمن. شوف تعليمات ADMIN_PASSWORD تحت." });
  }
  const account = requirePermission_(password, "viewData");
  if (!account) {
    return jsonOutput_({ status: "error", message: "Unauthorized" });
  }

  const formId = (e.parameter.form || "").trim();
  const cfg = getRegConfig_(formId);
  const requestedSheet = (e.parameter.sheet || "").trim();
  const sheetName = requestedSheet || cfg.activeSheetName;

  const sheet = findSheet_(sheetName);
  if (!sheet) {
    return jsonOutput_({ status: "error", message: `الشيت "${sheetName}" مش موجود.` });
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const rows = data.slice(1).filter(row => row.some(cell => String(cell).trim() !== ""));

  // Timestamps come back as Date objects from the Sheets API — stringify them
  // so the dashboard's JSON.parse-based date detection keeps working.
  const tsIdx = headers.indexOf("Timestamp");
  const serialized = rows.map(row => {
    const copy = row.slice();
    if (tsIdx > -1 && copy[tsIdx] instanceof Date) {
      copy[tsIdx] = Utilities.formatDate(copy[tsIdx], Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
    }
    return copy;
  });

  return jsonOutput_({
    status: "success",
    headers,
    rows: serialized,
    sheetName: sheet.getName(),
    activeSheetName: cfg.activeSheetName,
    role: account.role,
    permissions: account.permissions,
    accountName: account.name,
  });
}


// ---------------------------------------------------------------------------
// action=checkNid  (live duplicate check while typing — no password needed,
// it only ever answers true/false, never returns any personal data)
// ---------------------------------------------------------------------------

function handleCheckNid_(e) {
  const nationalId = (e.parameter.nationalId || "").trim();
  const formId = String(e.parameter.form || "").trim();
  if (!/^[0-9]{14}$/.test(nationalId)) {
    return jsonOutput_({ status: "success", exists: false });
  }
  return jsonOutput_({ status: "success", exists: isDuplicateNid_(nationalId, formId) });
}


// ---------------------------------------------------------------------------
// action=publicConfig  (public — the form calls this on load to know its
// title, its logo, and whether registration is currently open)
// ---------------------------------------------------------------------------

// Reads FIELD_CONFIG from Script Properties and fills in defaults for any
// field/key that's missing — so an install that never touched this Settings
// card behaves exactly like the original hardcoded validation.
function getFieldConfig_(formId) {
  const raw = PropertiesService.getScriptProperties().getProperty(propKey_("FIELD_CONFIG", formId));
  let stored = {};
  if (raw) {
    try { stored = JSON.parse(raw); } catch (e) { stored = {}; }
  }
  const result = {};
  Object.keys(ALL_BUILTIN_FIELDS).forEach((key, index) => {
    const s = stored[key] || {};
    result[key] = {
      enabled: typeof s.enabled === "boolean" ? s.enabled : true,
      required: typeof s.required === "boolean" ? s.required : ALL_BUILTIN_FIELDS[key].defaultRequired,
      // Drag-and-drop order within a section, from the "🧩 حقول الاستمارة"
      // card — defaults to registry order the first time (before anyone's
      // ever dragged anything), so ordering degrades gracefully to "however
      // they were defined" rather than a random/undefined order.
      order: typeof s.order === "number" ? s.order : index,
    };
  });
  return result;
}

// action=saveFieldConfig — admin toggles which fields are shown on the form
// and which are required (payload.fields), and/or replaces the whole list
// of admin-defined custom fields (payload.customFields), from the
// "🧩 حقول الاستمارة" settings card. Sending fields without customFields (or
// vice versa) only touches the one that was sent.
function handleSaveFieldConfig_(payload) {
  const formId = String(payload.formId || "").trim();
  const incoming = payload.fields || {};
  const sanitized = {};
  Object.keys(ALL_BUILTIN_FIELDS).forEach((key, index) => {
    const f = incoming[key] || {};
    sanitized[key] = {
      enabled: f.enabled !== false,   // default true unless explicitly turned off
      required: f.required === true,  // default false unless explicitly turned on
      order: typeof f.order === "number" ? f.order : index,
    };
  });
  PropertiesService.getScriptProperties().setProperty(propKey_("FIELD_CONFIG", formId), JSON.stringify(sanitized));

  let customFields = getCustomFields_(formId);
  if (Array.isArray(payload.customFields)) {
    const incomingKeys = payload.customFields.map(cf => cf && cf.key).filter(Boolean);
    customFields = payload.customFields.map(cf => sanitizeCustomField_(cf, incomingKeys)).filter(Boolean);
    // Custom fields also carry their own drag order (index in the array the
    // dashboard sends = the order the admin arranged them in).
    customFields.forEach((cf, i) => { cf.order = i; });
    saveCustomFields_(customFields, formId);
  }
  return jsonOutput_({ status: "success", fieldConfig: sanitized, customFields });
}

// ---------------------------------------------------------------------------
// Custom fields — admin-defined, unlimited, no code changes needed. Unlike
// the built-in fields above, these don't get their own sheet column each
// (see CUSTOM_FIELDS_HEADER) — their answers are stored as one JSON blob per
// registrant. Reference them in a certificate template as {{key}}, e.g.
// {{c_1a2b3c}} — the exact key is shown next to each custom field in the
// dashboard's "🧩 حقول الاستمارة" card after you add it.
// ---------------------------------------------------------------------------

function getCustomFields_(formId) {
  const raw = PropertiesService.getScriptProperties().getProperty(propKey_("CUSTOM_FIELDS", formId));
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function saveCustomFields_(fields, formId) {
  PropertiesService.getScriptProperties().setProperty(propKey_("CUSTOM_FIELDS", formId), JSON.stringify(fields));
}

const CUSTOM_FIELD_TYPES = ["text", "textarea", "number", "date", "select", "checkbox"];

// ---------------------------------------------------------------------------
// "Success screen" actions — the buttons shown to a registrant right after
// they submit the form (dashboard's "🎉 خيارات بعد التسجيل" card). Every
// possible option lives here as a TYPE; the dashboard just enables/disables
// individual entries and fills in the label/value — nothing here needs to
// change again when the org wants to add/remove a button, only the Settings
// tab does.
//
//   whatsapp_chat  — direct 1:1 WhatsApp chat with a number (e.g. HR/committee),
//                    opened with a pre-filled message containing the
//                    registrant's name + membership number.
//   whatsapp_group — join-link for a WhatsApp group/community
//                    (chat.whatsapp.com/... or whatsapp.com/channel/...).
//   telegram       — Telegram group/channel link (t.me/...).
//   facebook       — Facebook page/group link.
//   instagram      — Instagram profile link.
//   link           — any other custom link/button (Drive folder, survey,
//                     another website, ...).
const SUCCESS_ACTION_TYPES = ["whatsapp_chat", "whatsapp_group", "telegram", "facebook", "instagram", "link"];

function getSuccessActions_(formId) {
  const raw = PropertiesService.getScriptProperties().getProperty(propKey_("SUCCESS_ACTIONS", formId));
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function saveSuccessActions_(actions, formId) {
  PropertiesService.getScriptProperties().setProperty(propKey_("SUCCESS_ACTIONS", formId), JSON.stringify(actions));
}

// Forces an incoming {id, type, label, value, enabled} object into a trusted
// shape, same spirit as sanitizeCustomField_ above. Existing actions keep
// their original `id` (passed back by the dashboard) so re-saving the list
// after a small edit doesn't reshuffle anything.
function sanitizeSuccessAction_(a) {
  const label = String((a && a.label) || "").trim();
  const value = String((a && a.value) || "").trim();
  if (!label || !value) return null;
  const type = SUCCESS_ACTION_TYPES.indexOf(a && a.type) > -1 ? a.type : "link";
  const id = (a && a.id && /^sa_[a-z0-9]+$/.test(a.id)) ? a.id : "sa_" + Utilities.getUuid().replace(/-/g, "").slice(0, 8);
  return {
    id,
    type,
    label,
    value,
    enabled: a.enabled !== false,
  };
}

// Forces an incoming {label, type, options, required, enabled, key,
// dependsOnKey, dependsOnValue} object into a trusted shape. Existing custom
// fields keep their original `key` (passed back by the dashboard) so
// certificate placeholders referencing them don't silently break when you
// just tweak the label or required flag.
//
// dependsOnKey/dependsOnValue make this field conditional: if set, the form
// only shows (and only requires) this field once the OTHER field named
// dependsOnKey currently holds the value dependsOnValue — e.g. a "committee"
// field that only appears once "Are you a member?" is answered "Yes". Leave
// dependsOnKey empty for a normal, always-visible field (the default).
function sanitizeCustomField_(cf, validKeys) {
  const label = String((cf && cf.label) || "").trim();
  if (!label) return null;
  const type = CUSTOM_FIELD_TYPES.indexOf(cf.type) > -1 ? cf.type : "text";
  const key = (cf && cf.key && /^c_[a-z0-9]+$/.test(cf.key)) ? cf.key : "c_" + Utilities.getUuid().replace(/-/g, "").slice(0, 8);
  const out = {
    key,
    label,
    type,
    required: cf.required === true,
    enabled: cf.enabled !== false,
  };
  if (type === "select") {
    out.options = Array.isArray(cf.options)
      ? cf.options.map(String).map(s => s.trim()).filter(Boolean)
      : String(cf.options || "").split(",").map(s => s.trim()).filter(Boolean);
  }
  const dependsOnKey = String((cf && cf.dependsOnKey) || "").trim();
  // Only keep the dependency if it points at a real, different field — a
  // dangling reference (the other field got deleted) would otherwise hide
  // this one forever with no way to fix it from the UI.
  if (dependsOnKey && dependsOnKey !== key && (!validKeys || validKeys.indexOf(dependsOnKey) > -1)) {
    out.dependsOnKey = dependsOnKey;
    out.dependsOnValue = String((cf && cf.dependsOnValue) || "");
  }
  return out;
}

// ---------------------------------------------------------------------------
// AI chat assistant
// ---------------------------------------------------------------------------

function aiConfigured_(key) {
  return key && key.indexOf("PASTE_YOUR") === -1;
}

function checkAiQuota_() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  return Number(props.getProperty("AI_QUOTA_" + today) || "0") < AI_DAILY_LIMIT;
}

function bumpAiQuota_() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const key = "AI_QUOTA_" + today;
  props.setProperty(key, String(Number(props.getProperty(key) || "0") + 1));
}

// history: [{role:"user"|"bot", text}, ...] — the widget's own recent
// messages, so the assistant can follow a back-and-forth instead of
// answering each message in total isolation.
function callGemini_(systemPrompt, userMessage, history) {
  if (!aiConfigured_(GEMINI_API_KEY)) return null;
  try {
    const contents = (history || []).map(h => ({ role: h.role === "user" ? "user" : "model", parts: [{ text: h.text }] }));
    contents.push({ role: "user", parts: [{ text: userMessage }] });
    const res = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent",
      {
        method: "post",
        contentType: "application/json",
        muteHttpExceptions: true,
        // The newer "Auth key" format (AQ....) that AI Studio now issues by
        // default gets rejected with 401 ACCESS_TOKEN_TYPE_UNSUPPORTED when
        // sent as a ?key= query param — it needs to go in this header
        // instead. This also still works fine with an old-style AIzaSy...
        // key, so nothing breaks either way.
        headers: { "x-goog-api-key": GEMINI_API_KEY },
        payload: JSON.stringify({
          systemInstruction: { parts: [{ text: systemPrompt }] },
          contents,
          // gemini-3.5-flash-lite is the fast/cheap Gemini 3.x tier and has
          // thinking OFF by default (unlike 3.6 Flash, which thinks by
          // default and burns part of maxOutputTokens on it). That's exactly
          // right for short FAQ/stats replies, so no thinkingConfig needed —
          // just leave it at its default for max speed.
          generationConfig: {
            temperature: 0.4,
            maxOutputTokens: 600,
          },
        }),
      }
    );
    if (res.getResponseCode() !== 200) return null;
    const data = JSON.parse(res.getContentText());
    const parts = data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts;
    const text = parts && parts[0] && parts[0].text;
    return text ? text.trim() : null;
  } catch (e) {
    return null;
  }
}

function callGrok_(systemPrompt, userMessage, history) {
  if (!aiConfigured_(GROK_API_KEY)) return null;
  try {
    const messages = [{ role: "system", content: systemPrompt }];
    (history || []).forEach(h => messages.push({ role: h.role === "user" ? "user" : "assistant", content: h.text }));
    messages.push({ role: "user", content: userMessage });

    // GROK_API_KEY might be a real xAI key ("xai-...") or a Groq key
    // ("gsk_...") — those are two different companies with two different
    // endpoints/models. Auto-detect from the key's prefix so either one
    // just works, instead of silently failing when they don't match.
    const isGroq = GROK_API_KEY.indexOf("gsk_") === 0;
    const endpoint = isGroq ? "https://api.groq.com/openai/v1/chat/completions" : "https://api.x.ai/v1/chat/completions";
       // llama-3.3-70b-versatile was decommissioned by Groq on 16 Aug 2026 —
    // this is Groq's own recommended replacement.
    const model = isGroq ? "openai/gpt-oss-120b" : "grok-beta";

    const res = UrlFetchApp.fetch(endpoint, {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      headers: { Authorization: "Bearer " + GROK_API_KEY },
      payload: JSON.stringify({ model, messages, temperature: 0.4, max_tokens: 400 }),
    });
    if (res.getResponseCode() !== 200) return null;
    const data = JSON.parse(res.getContentText());
    const text = data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
    return text ? text.trim() : null;
  } catch (e) {
    return null;
  }
}

// Builds the "brain" the registrant-facing widget gets — only PUBLIC info
// (same data publicConfig already exposes), so it can accurately answer
// "when does registration close", "what fields do I need", "is there a
// WhatsApp group", etc. for THIS specific form, without ever seeing any
// registrant's actual data.
function buildFormAiPrompt_(cfg) {
  const fc = cfg.fieldConfig || {};
  const fieldLabels = Object.keys(ALL_BUILTIN_FIELDS)
    .filter(k => fc[k] && fc[k].enabled)
    .map(k => ALL_BUILTIN_FIELDS[k].label + (fc[k].required ? " (إجباري)" : " (اختياري)"));
  (cfg.customFields || []).forEach(c => fieldLabels.push(c.label + (c.required ? " (إجباري)" : " (اختياري)")));

  const phase = getRegPhase_(cfg);
  const phaseText = phase === "open" ? "التسجيل مفتوح دلوقتي"
    : phase === "before" ? `التسجيل لسه ماوصلش، هيفتح في ${cfg.startAt || "غير محدد"}`
    : `التسجيل قفل${cfg.endAt ? " من " + cfg.endAt : ""}`;

  const actionsText = (cfg.successActions || [])
    .filter(a => a.enabled !== false)
    .map(a => `- ${a.label}`).join("\n") || "(مفيش)";

  return (
    `انت مساعد ودود بيرد على أسئلة الناس اللي بتحاول تسجل في استمارة "${cfg.formTitle || "التسجيل"}" ` +
    `بتاعة أسرة صناع الحياة (مبادرة طلابية تطوعية في الجامعات المصرية).\n\n` +
    `معلومات عن الاستمارة دي:\n` +
    `- حالة التسجيل: ${phaseText}\n` +
    `- الحقول المطلوبة/الموجودة في الاستمارة: ${fieldLabels.join("، ") || "الاسم والرقم القومي بس"}\n` +
    `- بعد التسجيل، طرق التواصل المتاحة: \n${actionsText}\n\n` +
    `تعليمات مهمة:\n` +
    `- رد بالعربي العامي المصري، بإيجاز ووضوح (2-4 جمل عادةً).\n` +
    `- لو حد سأل سؤال مش عن التسجيل ده خالص (زي أسئلة عامة أو تقنية بعيدة)، رد بلطف إنك هنا للمساعدة في التسجيل بس.\n` +
    `- لو حد سأل عن بياناته الشخصية أو بيانات حد تاني، وضّح إنك معندكش وصول لأي بيانات مسجلين، واقترح يتواصل مع فريق اللجنة.\n` +
    `- ماتخترعش معلومات مش موجودة قدامك — لو مش متأكد، قول إنك مش متأكد واقترح التواصل المباشر.`
  );
}

// The actual "how do I..." steps the AI is allowed to hand back verbatim —
// kept here (not invented by the model) so answers match the real UI.
// Update this alongside any dashboard UI changes.
const DASHBOARD_HOWTO_GUIDE =
  `- إزاي أرفع شعار للفورم؟ روح تبويب "⚙️ الإعدادات" ▸ كارت "صورة الاستمارة (الشعار)" ▸ دوس اختار ملف صورة من جهازك ▸ دوس زرار "رفع الصورة". لو عايز ترجع للشعار الافتراضي، دوس "استخدام الشعار الافتراضي" جنبه.\n` +
  `- إزاي أضيف فورم جديد؟ فوق الصفحة جنب اختيار الفورم الحالي، دوس زرار "➕ فورم جديد"، هيطلب منك اسم الفورم واكتبه ودوس موافق — الفورم الجديد هيبقى له لينك وإعدادات وشيت منفصلين تمامًا عن أي فورم تاني.\n` +
  `- إزاي أوافق على حساب جوجل جديد؟ روح تبويب "⚙️ الإعدادات" ▸ كارت "👥 حسابات الدخول" ▸ هتلاقي أي طلب دخول جديد ظاهر بعلامة "(بانتظار الموافقة)" في الجدول ▸ جنبه زرار موافقة ورفض، دوس "موافقة" وحدد له الصلاحيات المناسبة (إعدادات عامة / حقول الاستمارة / الشهادات / إدارة الحسابات) وبعدين احفظ. (بديل: لو معاه إيميل مسجل كـ admin بصلاحية "إدارة الحسابات"، بيوصله إيميل فيه لينك موافقة/رفض مباشر).\n` +
  `- إزاي أغيّر إعدادات عامة (العنوان، ميعاد التسجيل، الدورات)؟ تبويب "⚙️ الإعدادات" ▸ أول كارت في الصفحة ▸ عدّل الحقول ودوس "حفظ الإعدادات".\n` +
  `- إزاي أضيف/أشيل حقول من الفورم؟ تبويب "⚙️ الإعدادات" ▸ كارت "🧩 حقول الاستمارة" ▸ فعّل/عطّل أو اجعل الحقل إجباري من هناك.\n` +
  `- إزاي أبعت الشهادات؟ تبويب "⚙️ الإعدادات" ▸ كارت الشهادات ▸ ارفع تيمبلت الشهادة (docx) ▸ ابعت شهادة تجريبية للتأكد ▸ بعدين ابعت الشهادات فعليًا (تلقائي أو دفعة واحدة حسب الإعداد).\n` +
  `- إزاي أعمل تشخيص لو حاجة مش شغالة (زي الإيميلات)؟ تبويب "⚙️ الإعدادات" ▸ زرار "🩺 تشخيص" بيوريك تقرير كامل (حالة الإيميل، الصلاحيات، الربط بجوجل درايف...).`;

function buildDashboardAiPrompt_(cfg, stats) {
  return (
    `انت مساعد بيساعد الأدمن (مسؤول لوحة تحكم أسرة صناع الحياة) يفهم بيانات ولوحة تحكم فورم "${cfg.formTitle || "بدون عنوان"}" ` +
    `ويستخدم الداشبورد صح.\n\n` +
    `إحصائيات إجمالية عن الفورم ده (الشيت الحالي: ${cfg.activeSheetName}):\n` +
    `- إجمالي المسجلين: ${stats.total}\n` +
    `- نسبة الإناث: ${stats.femalePct}%\n` +
    `- نسبة الخريجين: ${stats.graduatePct}%\n` +
    `- نسبة اللي بيشتغلوا حاليًا: ${stats.employedPct}%\n` +
    `- تسجيلات آخر 7 أيام: ${stats.last7Days}\n` +
    `- أكتر 3 صفات/كليات تكرارًا: ${stats.topBreakdown}\n\n` +
    `دليل استخدام الداشبورد (المصدر الوحيد اللي تقدر تجاوب منه على أسئلة "إزاي أعمل كذا" — استخدم نفس الخطوات دي بالظبط، متخترعش خطوات تانية):\n` +
    DASHBOARD_HOWTO_GUIDE + `\n\n` +
    `تعليمات مهمة:\n` +
    `- رد بالعربي العامي المصري، بإيجاز.\n` +
    `- استخدم الأرقام دي بس للإجابة عن أسئلة إحصائية — ماعندكش وصول لأي بيانات شخصية لأي مسجل (اسم، رقم قومي، تليفون، إيميل) خالص، فلو سأل عن حد بعينه وضّح إنه يشوف الجدول نفسه في تبويب "📋 الجدول".\n` +
    `- لو سأل "إزاي أعمل كذا" وموجود في دليل الاستخدام فوق، جاوب بنفس خطواته بالظبط. لو مش موجود في الدليل، قول بصراحة إنك مش متأكد من الخطوة دي واقترح إنه يدوّر في تبويب "⚙️ الإعدادات" أو يسأل فريق التقنية.\n` +
    `- ماتخترعش أرقام أو معلومات أو خطوات مش معطاة لك.`
  );
}

// Only counts/percentages/breakdowns — deliberately never returns a single
// row of actual registrant data (see buildDashboardAiPrompt_ above).
function computeAggregateStats_(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const rows = data.slice(1).filter(r => r.some(c => String(c).trim() !== ""));
  const col = (name) => headers.indexOf(name);

  const genderCol = col("Gender"), gradCol = col("Graduate"), jobCol = col("Has Job"),
    facultyCol = col("Faculty"), tsCol = col("Timestamp");

  const total = rows.length;
  const pct = (n) => total ? Math.round((n / total) * 100) : 0;

  const femaleCount = genderCol > -1 ? rows.filter(r => /أنث|بنت|female/i.test(String(r[genderCol]))).length : 0;
  const gradCount = gradCol > -1 ? rows.filter(r => /^(نعم|yes|true)/i.test(String(r[gradCol]))).length : 0;
  const jobCount = jobCol > -1 ? rows.filter(r => /^(نعم|yes|true)/i.test(String(r[jobCol]))).length : 0;

  const weekAgo = Date.now() - 7 * 24 * 60 * 60 * 1000;
  const last7Days = tsCol > -1 ? rows.filter(r => {
    const t = r[tsCol] instanceof Date ? r[tsCol].getTime() : new Date(r[tsCol]).getTime();
    return !isNaN(t) && t >= weekAgo;
  }).length : 0;

  const facultyCounts = {};
  if (facultyCol > -1) rows.forEach(r => {
    const v = String(r[facultyCol]).trim();
    if (v) facultyCounts[v] = (facultyCounts[v] || 0) + 1;
  });
  const topBreakdown = Object.entries(facultyCounts).sort((a, b) => b[1] - a[1]).slice(0, 3)
    .map(([k, v]) => `${k} (${v})`).join("، ") || "غير متاح";

  return { total, femalePct: pct(femaleCount), graduatePct: pct(gradCount), employedPct: pct(jobCount), last7Days, topBreakdown };
}

// action=chatAI (POST) — payload: {context: "form"|"dashboard", formId,
// message, history, password (dashboard only)}.
function handleChatAI_(payload) {
  const context = payload.context === "dashboard" ? "dashboard" : "form";
  const message = String(payload.message || "").trim().slice(0, 800);
  if (!message) return jsonOutput_({ status: "error", message: "اكتب سؤالك الأول." });

  if (context === "dashboard") {
    const account = requirePermission_(payload.password || "", "viewData");
    if (!account) return jsonOutput_({ status: "error", message: "Unauthorized" });
  }

  if (!aiConfigured_(GEMINI_API_KEY) && !aiConfigured_(GROK_API_KEY)) {
    return jsonOutput_({ status: "error", message: "المساعد الذكي لسه مش متفعّل — استخدم الأسئلة الجاهزة تحت." });
  }
  if (!checkAiQuota_()) {
    return jsonOutput_({ status: "error", message: "المساعد وصل للحد اليومي من الأسئلة — جرب تاني بكرة، أو استخدم الأسئلة الجاهزة." });
  }

  const formId = String(payload.formId || "").trim();
  const cfg = getRegConfig_(formId);
  let systemPrompt;
  if (context === "dashboard") {
    const sheet = findSheet_(cfg.activeSheetName);
    const stats = sheet ? computeAggregateStats_(sheet) : { total: 0, femalePct: 0, graduatePct: 0, employedPct: 0, last7Days: 0, topBreakdown: "غير متاح" };
    systemPrompt = buildDashboardAiPrompt_(cfg, stats);
  } else {
    systemPrompt = buildFormAiPrompt_(cfg);
  }

    const history = Array.isArray(payload.history) ? payload.history.slice(-8) : [];
  // Groq tried FIRST — Gemini is currently blocked by a known Google-side
  // bug affecting "AQ." auth keys (401 ACCESS_TOKEN_TYPE_UNSUPPORTED on
  // generateContent, widely reported, no fix yet as of Sept 2026). Swap
  // this back to Gemini-first once Google resolves it.
  let reply = callGrok_(systemPrompt, message, history);
  let provider = "grok";
  if (!reply) { reply = callGemini_(systemPrompt, message, history); provider = "gemini"; }
  if (!reply) {
    return jsonOutput_({ status: "error", message: "المساعد مش متاح دلوقتي — جرب تاني كمان شوية، أو استخدم الأسئلة الجاهزة." });
  }
  bumpAiQuota_();
  return jsonOutput_({ status: "success", reply, provider });
}

function handlePublicConfig_(e) {
  const formId = String((e && e.parameter && e.parameter.form) || "").trim();
  const cfg = getRegConfig_(formId);
  const phase = getRegPhase_(cfg); // "before" | "open" | "closed"
  return jsonOutput_({
    status: "success",
    formTitle: cfg.formTitle,
    logoUrl: cfg.logoUrl,
    startAt: cfg.startAt,
    endAt: cfg.endAt,
    phase,
    fieldConfig: cfg.fieldConfig,
    fieldDefs: cfg.fieldDefs,
    fieldSections: cfg.fieldSections,
    customFields: cfg.customFields,
    // Public form only ever needs the ENABLED buttons — disabled ones stay
    // hidden from anyone inspecting the public endpoint, not just from the UI.
    successActions: (cfg.successActions || []).filter(a => a.enabled !== false),
    serverNow: new Date().toISOString(),
  });
}

// action=getConfig  (owner only — prefills the Settings tab in the dashboard)
function handleGetConfig_(e) {
  if (!requirePermission_(e.parameter.password || "", "manageSettings")) {
    return jsonOutput_({ status: "error", message: "Unauthorized" });
  }
  const formId = String(e.parameter.form || "").trim();
  return jsonOutput_({ status: "success", config: getRegConfig_(formId) });
}

// Which sheets were EXPLICITLY created for a given form (via startNewCycle_
// or getSheet_'s auto-create fallback) — see addFormSheetName_ below. This
// is what handleListCycles_ uses to know which sheets belong to which form,
// instead of guessing from sheetBaseName string-matching (which broke for
// legacy sheets whose names don't follow the "base"/"base 2"/"base 3"
// pattern, e.g. a pre-existing "Responses_1" tab).
function getFormSheetNames_(formId) {
  const raw = PropertiesService.getScriptProperties().getProperty(propKey_("CYCLE_SHEETS", formId));
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function addFormSheetName_(formId, name) {
  const props = PropertiesService.getScriptProperties();
  const key = propKey_("CYCLE_SHEETS", formId);
  const names = getFormSheetNames_(formId);
  if (names.indexOf(name) === -1) {
    names.push(name);
    props.setProperty(key, JSON.stringify(names));
  }
}

// action=listCycles  (anyone logged in — lets the dashboard switch which
// sheet/cycle's data it's showing; viewing a cycle only needs viewData).
//
// The default form ("") shows every sheet NOT explicitly claimed by another
// (non-default) form — this is what makes old, pre-multi-form sheets (e.g.
// "Responses_1", created before this ownership tracking existed) show up
// correctly instead of disappearing, since nothing else has claimed them.
// A specific (non-default) form only ever shows sheets it created itself.
function handleListCycles_(e) {
  if (!requirePermission_(e.parameter.password || "", "viewData")) {
    return jsonOutput_({ status: "error", message: "Unauthorized" });
  }
  const formId = String(e.parameter.form || "").trim();
  const cfg = getRegConfig_(formId);
  const ss = getSpreadsheet_();

  const claimedByOthers = new Set();
  listAllForms_(true).forEach(f => {
    if (f.id === formId) return;
    getFormSheetNames_(f.id).forEach(n => claimedByOthers.add(n));
  });

  let names = formId
    ? getFormSheetNames_(formId)
    : ss.getSheets().map(s => s.getName()).filter(n => n !== ACTIVITY_LOG_SHEET_NAME && !claimedByOthers.has(n));

  // Safety net: never hide the sheet this form is CURRENTLY writing into,
  // even if it's somehow missing from the tracked list.
  if (cfg.activeSheetName && names.indexOf(cfg.activeSheetName) === -1) names.push(cfg.activeSheetName);

  const sheets = names
    .map(n => ss.getSheetByName(n))
    .filter(Boolean)
    .map(sh => ({
      name: sh.getName(),
      rows: Math.max(0, sh.getLastRow() - 1),
      isActive: sh.getName() === cfg.activeSheetName,
    }));
  return jsonOutput_({ status: "success", sheets, activeSheetName: cfg.activeSheetName });
}

// action=diagnostics  (owner only) — runs a handful of read-only checks and
// reports exactly what's working and what isn't, instead of you having to
// guess from a generic "error" message. Open ?action=diagnostics&password=...
// directly in a browser, or use the "🩺 تشخيص" button in the dashboard.
function handleDiagnostics_(e) {
  if (!requirePermission_(e.parameter.password || "", "manageSettings")) {
    return jsonOutput_({ status: "error", message: "Unauthorized" });
  }

  const report = {};

  // 1) Deployed code version marker — bump CODE_VERSION at the top of this
  //    file (or just eyeball this string) to confirm the web app is actually
  //    running the version you last pasted in, and not a stale deployment.
  report.codeVersion = typeof CODE_VERSION !== "undefined" ? CODE_VERSION : "(CODE_VERSION not set)";

  // 2) Admin accounts configured?
  report.adminAccounts = getAdminAccounts_().map(a => ({ name: a.name, role: a.role, permissions: a.permissions }));

  // 3) Script timezone (affects {{date}} / {{registrationDate}} on certificates)
  try { report.scriptTimeZone = Session.getScriptTimeZone(); }
  catch (err) { report.scriptTimeZone = "خطأ: " + String(err); }

  // 4) Active sheet reachable? (default form — diagnostics is a global
  //    health-check; report.otherForms below lists any additional forms)
  try {
    const cfg = getRegConfig_("");
    report.activeSheetName = cfg.activeSheetName;
    const sheet = findSheet_(cfg.activeSheetName);
    report.activeSheetFound = !!sheet;
    report.fieldConfig = cfg.fieldConfig;
    report.sendCertAuto = cfg.sendCertAuto;
    report.certTemplateReady = cfg.certTemplateReady;
  } catch (err) {
    report.sheetCheckError = String(err);
  }

  // 5) MailApp — remaining daily quota. If this is 0, that's exactly why
  //    emails (confirmation AND certificate) silently stop sending — Gmail
  //    accounts get ~100/day, more with Google Workspace.
  try { report.mailRemainingQuota = MailApp.getRemainingDailyQuota(); }
  catch (err) { report.mailQuotaError = String(err); }

  // 6) Drive API advanced service — required for certificate template
  //    upload/conversion. If this errors, go enable it: Apps Script editor ▸
  //    Services ▸ + ▸ Drive API (see BACKEND_SETUP_STEPS.md section 11-أ).
  try {
    if (typeof Drive === "undefined") {
      report.driveApiAdvancedService = "غير مفعّلة — روح فعّلها من Services جوه محرر الكود (Drive API).";
    } else {
      Drive.About.get({ fields: "user" }); // v3 requires an explicit `fields` param — trivial read-only call just to confirm the service works
      report.driveApiAdvancedService = "شغالة ✓";
    }
  } catch (err) {
    report.driveApiAdvancedService = "مفعّلة بس بترمي خطأ: " + String(err);
  }

  // 7) DocumentApp OAuth scope — separate from "Drive API advanced service"
  //    above. generateCertificatePdf_() calls DocumentApp.openById() to fill
  //    in the template, which needs the documents scope. Enabling the Drive
  //    advanced service switches this project to an EXPLICIT scope list in
  //    its manifest (appsscript.json) instead of auto-detecting scopes from
  //    your code — so the documents scope can go missing even though
  //    everything else (template upload, mail quota) reports fine. If this
  //    fails, see "لو الشهادة التجريبية بترجع Exception: ليس لديك إذن..." in
  //    BACKEND_SETUP_STEPS.md.
  try {
    const tempDoc = DocumentApp.create("dys-diagnostics-scope-check-temp");
    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
    report.documentAppScope = "شغالة ✓";
  } catch (err) {
    report.documentAppScope = "بترمي خطأ: " + String(err) +
      " — لازم تضيف https://www.googleapis.com/auth/documents لـ oauthScopes في appsscript.json وتعيد الموافقة (شوف قسم الشهادات في BACKEND_SETUP_STEPS.md).";
  }

  // 8) Firestore mirror — optional and additive (see pushToFirestore_
  //    above); a missing setup here never blocks real registrations.
  try {
    const props = PropertiesService.getScriptProperties();
    const hasCreds = !!(props.getProperty("FIRESTORE_PROJECT_ID") && props.getProperty("FIRESTORE_CLIENT_EMAIL") && props.getProperty("FIRESTORE_PRIVATE_KEY"));
    if (!hasCreds) {
      report.firestoreMirror = "غير مُعدّ (اختياري) — شوف قسم فايربيز في BACKEND_SETUP_STEPS.md لو عايز تفعّله.";
    } else if (typeof FirestoreApp === "undefined") {
      report.firestoreMirror = "الإعدادات موجودة بس مكتبة FirestoreApp مش مضافة للمشروع (Libraries).";
    } else {
      const firestore = FirestoreApp.getFirestore(
        props.getProperty("FIRESTORE_CLIENT_EMAIL"),
        props.getProperty("FIRESTORE_PRIVATE_KEY"),
        props.getProperty("FIRESTORE_PROJECT_ID")
      );
      firestore.createDocument("registrations/_diagnostics_check", { checkedAt: new Date().toISOString() });
      report.firestoreMirror = "شغالة ✓";
    }
  } catch (err) {
    report.firestoreMirror = "بترمي خطأ: " + String(err);
  }

  return jsonOutput_({ status: "success", report });
}


// ---------------------------------------------------------------------------
// doPost  (form submission)
// ---------------------------------------------------------------------------

// Uploads a registrant's personal photo (payload.photoBase64, a data: URI
// or raw base64 string sent by the form after client-side resizing) to a
// dedicated "أسرة صناع الحياة - صور المسجلين" Drive folder, and returns a public
// view-only URL — or "" on any failure (a broken photo upload must never
// block someone's registration, so this is best-effort and swallows errors).
function uploadRegistrationPhoto_(photoBase64) {
  if (!photoBase64) return "";
  try {
    const raw = String(photoBase64);
    const commaIdx = raw.indexOf(",");
    const base64 = commaIdx > -1 && raw.slice(0, commaIdx).indexOf("base64") > -1 ? raw.slice(commaIdx + 1) : raw;
    const mimeMatch = raw.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64/);
    const mimeType = mimeMatch ? mimeMatch[1] : "image/jpeg";
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mimeType, "photo-" + new Date().getTime() + ".jpg");

    const folder = getOrCreatePhotosFolder_();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/uc?export=view&id=" + file.getId();
  } catch (err) {
    return "";
  }
}

function getOrCreatePhotosFolder_() {
  const FOLDER_NAME = "أسرة صناع الحياة - صور المسجلين";
  const existing = DriveApp.getFoldersByName(FOLDER_NAME);
  if (existing.hasNext()) return existing.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

// ---------------------------------------------------------------------------
// Firestore mirror (optional, additive) — every successful registration is
// ALSO written to a Firestore "registrations" collection, in addition to
// (never instead of) the Google Sheet. Needs one-time setup:
//   1. Add the FirestoreApp library to this project: Apps Script editor ▸
//      Libraries (+) ▸ paste script ID
//      1VUSl4b1r1eoNcRWotZM3e87ygkxvXltOgyDZhixqncz9lQ3MjfT1iKFw ▸ Add ▸
//      pick the latest version ▸ keep identifier "FirestoreApp".
//   2. Set three Script Properties from a Firebase service account JSON key:
//      FIRESTORE_PROJECT_ID, FIRESTORE_CLIENT_EMAIL, FIRESTORE_PRIVATE_KEY.
//      See the "فايربيز" section in BACKEND_SETUP_STEPS.md for exact steps.
// If any of that isn't set up (or the library isn't added yet), this
// silently does nothing — a missing/broken Firestore mirror must NEVER
// block or fail an actual registration, the Sheet write already happened.
function pushToFirestore_(membershipNo, rowValues) {
  try {
    const props = PropertiesService.getScriptProperties();
    const projectId = props.getProperty("FIRESTORE_PROJECT_ID");
    const clientEmail = props.getProperty("FIRESTORE_CLIENT_EMAIL");
    const privateKey = props.getProperty("FIRESTORE_PRIVATE_KEY");
    if (!projectId || !clientEmail || !privateKey) return false;
    if (typeof FirestoreApp === "undefined") return false; // library not added yet

    const firestore = FirestoreApp.getFirestore(clientEmail, privateKey, projectId);
    const data = {};
    HEADERS.forEach((h, i) => {
      const v = rowValues[i];
      data[h] = v instanceof Date ? v.toISOString() : String(v == null ? "" : v);
    });
    firestore.createDocument("registrations/" + membershipNo, data);
    return true;
  } catch (err) {
    return false;
  }
}

function handleSubmit_(payload) {
  // ---- 1) anti-bot checks (mirrors the frontend's own honeypot + timing check) ----
  if (payload.website) { // honeypot field — a real user never fills this in
    return jsonOutput_({ status: "error", message: "Rejected" });
  }
  const elapsed = Date.now() - Number(payload.loadedAt || 0);
  if (!payload.loadedAt || isNaN(elapsed) || elapsed < MIN_FILL_MS) {
    return jsonOutput_({ status: "error", message: "Submitted too fast" });
  }

  const formId = String(payload.formId || "").trim();

  // ---- 2) is registration currently open? ----
  const cfg = getRegConfig_(formId);
  const phase = getRegPhase_(cfg);
  if (phase !== "open") {
    return jsonOutput_({
      status: phase === "before" ? "not_started" : "closed",
      startAt: cfg.startAt,
      endAt: cfg.endAt,
    });
  }

  // ---- 3) validate every field server-side — never trust the client alone ----
  const errors = validatePayload_(payload, formId);
  if (errors.length) {
    return jsonOutput_({ status: "error", message: errors.join(" | ") });
  }

  const nationalId = String(payload.nationalId).trim();

  // ---- 4) duplicate guard (uses a lock so two near-simultaneous submits
  //         of the same ID can't both slip through) ----
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (isDuplicateNid_(nationalId, formId)) {
      return jsonOutput_({ status: "duplicate" });
    }

    // ---- 5) generate membership number (or reuse an existing one, if this
    //         same person already registered before — see
    //         findExistingMembershipNo_) + append the row ----
    const membershipNo = findExistingMembershipNo_(nationalId) || generateMembershipNumber_(cfg.membershipPrefix);
    const fc = getFieldConfig_(formId);
    if (fc.photo && fc.photo.enabled && payload.photoBase64) {
      payload.photoUrl = uploadRegistrationPhoto_(payload.photoBase64);
    }
    const sheet = getSheet_(formId);
    const rowValues = buildRow_(payload, membershipNo);
    sheet.appendRow(rowValues);
    pushToFirestore_(membershipNo, rowValues); // best-effort mirror — never blocks registration

    // ---- 6) confirmation email (best-effort — never fails the submission) ----
    let emailSent = false;
    const validEmail = payload.email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(payload.email.trim());
    if (validEmail) {
      emailSent = sendConfirmationEmail_(payload, membershipNo, cfg);
    }

    // ---- 7) certificate email (best-effort, only if auto-send is turned on
    //         from the dashboard AND a certificate template is uploaded) ----
    let certificateSent = false;
    if (validEmail && cfg.sendCertAuto && cfg.certTemplateReady) {
      certificateSent = sendCertificateEmail_(payload, membershipNo, formId);
    }

    return jsonOutput_({ status: "success", membershipNo, emailSent, certificateSent });
  } finally {
    lock.releaseLock();
  }
}


// ---------------------------------------------------------------------------
// Validation — mirrors the frontend's validators[] exactly (dys_form.html)
// ---------------------------------------------------------------------------

function validatePayload_(p, formId) {
  const errors = [];
  const str = (v) => String(v || "").trim();
  const fc = getFieldConfig_(formId);
  const isOn = (key) => fc[key] ? fc[key].enabled : true;
  const isReq = (key) => fc[key] ? fc[key].required : true;

  // Name + National ID are always required — the whole duplicate-check and
  // membership system depends on them, so they're not part of TOGGLEABLE_FIELDS.
  if (str(p.name).length < 3) errors.push("name");
  if (!isValidEgyptianNationalId_(str(p.nationalId))) errors.push("nationalId");

  if (isOn("age")) {
    const v = str(p.age);
    if (isReq("age") || v !== "") {
      const age = Number(v);
      if (!/^[0-9]{1,2}$/.test(v) || age <= 0 || age >= 100) errors.push("age");
    }
  }

  if (isOn("gender")) {
    if (isReq("gender") || p.gender) {
      if (!["Male", "Female"].includes(p.gender)) errors.push("gender");
    }
  }

  if (isOn("phone")) {
    const v = str(p.phone);
    if (isReq("phone") || v !== "") {
      if (!/^01[0125][0-9]{8}$/.test(v)) errors.push("phone");
    }
  }

  if (isOn("whatsapp")) {
    const v = str(p.whatsapp);
    if (isReq("whatsapp") || v !== "") {
      if (!/^01[0125][0-9]{8}$/.test(v)) errors.push("whatsapp");
    }
  }

  if (isOn("email")) {
    const v = str(p.email);
    if (isReq("email") && v === "") errors.push("email");
    else if (v && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) errors.push("email");
  }

  if (isOn("faculty")) {
    const v = str(p.faculty);
    if (isReq("faculty") || v !== "") {
      if (v.length < 2) errors.push("faculty");
    }
  }

  if (isOn("graduate")) {
    if (isReq("graduate") || p.graduate) {
      if (!["Yes", "No"].includes(p.graduate)) errors.push("graduate");
    }
  }

  if (isOn("committee")) {
    const v = str(p.committee);
    if (isReq("committee") || v !== "") {
      if (v.length < 1) errors.push("committee");
    }
  }

  if (isOn("hasJob")) {
    if (isReq("hasJob") || p.hasJob) {
      if (!["Yes", "No"].includes(p.hasJob)) errors.push("hasJob");
    }
  }

  // Generic validation for the newer EXTRA_FIELDS (address, birthDate,
  // maritalStatus, ...) — unlike the hand-written checks above, these are
  // all driven purely by `type` (see EXTRA_FIELDS at the top of the file).
  Object.keys(EXTRA_FIELDS).forEach(key => {
    if (!isOn(key)) return;
    const def = EXTRA_FIELDS[key];
    const v = str(p[key]);
    if (v === "") {
      if (isReq(key)) errors.push(key);
      return;
    }
    if (def.type === "select" && def.options && def.options.indexOf(p[key]) === -1) errors.push(key);
    if (def.type === "date" && !/^\d{4}-\d{2}-\d{2}$/.test(v)) errors.push(key);
    if (def.type === "checkbox" && ["Yes", "No"].indexOf(v) === -1) errors.push(key);
  });

  // Admin-defined custom fields — payload.customFields is {key: value, ...}.
  const customFieldsByKey_ = {};
  const customFieldsList_ = getCustomFields_(formId);
  customFieldsList_.forEach(cf => { customFieldsByKey_[cf.key] = cf; });
  customFieldsList_.forEach(cf => {
    if (!cf.enabled) return;
    // Conditional field (dependsOnKey/dependsOnValue) whose condition isn't
    // met right now — the form hides it in that case, so it can't be
    // required either, regardless of what the admin set for `required`.
    if (cf.dependsOnKey) {
      const parentDef = customFieldsByKey_[cf.dependsOnKey];
      const parentVal = parentDef ? str((p.customFields || {})[cf.dependsOnKey]) : "";
      if (parentVal !== cf.dependsOnValue) return;
    }
    const v = str((p.customFields || {})[cf.key]);
    if (v === "") {
      if (cf.required) errors.push(cf.key);
      return;
    }
    if (cf.type === "select" && cf.options && cf.options.indexOf(v) === -1) errors.push(cf.key);
  });

  // Photo — just a presence check here (payload.photoBase64 non-empty).
  // The actual upload/failure handling happens later in handleSubmit_ and
  // never blocks the registration even if it fails.
  if (isOn("photo") && isReq("photo") && !p.photoBase64) {
    errors.push("photo");
  }

  // Video — unlike photo, this is uploaded straight from the browser to
  // Cloudinary BEFORE the form is even submitted (see dys_form.html), so by
  // the time it gets here p.videoUrl is either empty or already a real
  // Cloudinary URL. Just a presence + basic sanity check.
  if (isOn("video")) {
    const v = str(p.videoUrl);
    if (v === "") {
      if (isReq("video")) errors.push("video");
    } else if (!/^https:\/\//.test(v)) {
      errors.push("video");
    }
  }

  return errors;
}

// Same structural check the frontend does client-side (century digit, valid
// month/day, valid governorate code, plausible age) — see extractNationalIdInfo()
// in dys_form.html for the reference implementation this mirrors.
function isValidEgyptianNationalId_(id) {
  if (!/^[0-9]{14}$/.test(id)) return false;

  const centuryDigit = id[0];
  const centuryMap = { "2": 1900, "3": 2000 };
  if (!(centuryDigit in centuryMap)) return false;

  const yy = parseInt(id.slice(1, 3), 10);
  const mm = parseInt(id.slice(3, 5), 10);
  const dd = parseInt(id.slice(5, 7), 10);
  const govCode = id.slice(7, 9);

  const validGovCodes = ["01","02","03","04","11","12","13","14","15","16","17","18",
    "19","21","22","23","24","25","26","27","28","29","31","32","33","34","35","88"];
  if (!validGovCodes.includes(govCode)) return false;

  if (mm < 1 || mm > 12) return false;
  const birthYear = centuryMap[centuryDigit] + yy;
  const daysInMonth = new Date(birthYear, mm, 0).getDate();
  if (dd < 1 || dd > daysInMonth) return false;

  const now = new Date();
  let age = now.getFullYear() - birthYear;
  if (now.getMonth() + 1 < mm || (now.getMonth() + 1 === mm && now.getDate() < dd)) age--;
  if (age < 0 || age > 110) return false;

  return true;
}


// ---------------------------------------------------------------------------
// Registration window / cycle config
// ---------------------------------------------------------------------------
// Stored entirely in Script Properties (not in this source file), so the
// dashboard's Settings tab can change everything without ever touching code.
//
//   FORM_TITLE        — shown as the form's page title (falls back to the
//                        form's own default text if left empty)
//   SHEET_BASE_NAME    — the "base" tab name each new cycle is built from
//   REG_START/REG_END  — ISO datetime strings, or "" for "no limit"
//   ACTIVE_SHEET_NAME   — the tab the form is CURRENTLY writing into
//   CYCLE_NUMBER        — how many cycles/sheets have been created so far
//   LOGO_URL/LOGO_FILE_ID — the uploaded logo image, stored in Drive
//
// ---------------------------------------------------------------------------
// Multi-form support ("عدة فورمات في نفس الوقت")
// ---------------------------------------------------------------------------
// Everything above is stored under a plain key like "FORM_TITLE". To let
// several independent registration forms run at the same time from this one
// backend/spreadsheet (each with its own link, title, dates, fields,
// success-screen buttons, cycles/sheets, and optional certificate template),
// every one of those keys gets namespaced per form via propKey_() below —
// e.g. "FORM_TITLE__f_a1b2c3" instead of "FORM_TITLE".
//
// The ORIGINAL, pre-multi-form data is form id "" (empty string) — so an
// existing deployment with real registrations keeps working with ZERO
// migration: its old un-namespaced keys ("FORM_TITLE", "ACTIVE_SHEET_NAME",
// ...) are exactly what formId "" resolves to. dys_form.html with no
// `?form=` in its URL (the original link) is that same default form.
//
// New forms are created from the dashboard's form switcher ("➕ فورم جديد")
// — see handleAddForm_ — which assigns a short random id and its own
// sheetBaseName, then everything else (settings, fields, success actions,
// cycles) is scoped to that id automatically the moment you're managing it.
const FORMS_REGISTRY_KEY = "FORMS_REGISTRY";

function propKey_(base, formId) {
  return formId ? `${base}__${formId}` : base;
}

// Registry of every EXTRA form (beyond the default "" one, which always
// exists implicitly and is never stored in this list). Each entry:
// {id, name, createdAt}.
function getFormsRegistry_() {
  const raw = PropertiesService.getScriptProperties().getProperty(FORMS_REGISTRY_KEY);
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function saveFormsRegistry_(forms) {
  PropertiesService.getScriptProperties().setProperty(FORMS_REGISTRY_KEY, JSON.stringify(forms));
}

// Full list including the default form, in the shape the dashboard's form
// switcher needs — default form's name comes from its OWN formTitle setting
// (falls back to a generic label if it was never given a title).
// includeArchived=false (the default) hides archived forms — the default
// form can never be archived, so it's always included.
function listAllForms_(includeArchived) {
  const defaultCfg = getRegConfig_("");
  const defaultEntry = { id: "", name: defaultCfg.formTitle || "الفورم الأساسي", isDefault: true, archived: false };
  const extra = getFormsRegistry_().map(f => Object.assign({ isDefault: false, archived: false }, f));
  return [defaultEntry].concat(includeArchived ? extra : extra.filter(f => !f.archived));
}

// action=listForms (GET, viewData) — anyone logged in can see the list (the
// dashboard needs it just to populate the "أي فورم" switcher), only
// manageSettings can actually ADD/ARCHIVE one (see handleAddForm_/
// handleArchiveForm_ below). ?includeArchived=1 also returns archived forms
// (used by the "اعرض المؤرشف" checkbox in the dashboard).
function handleListForms_(e) {
  if (!requirePermission_(e.parameter.password || "", "viewData")) {
    return jsonOutput_({ status: "error", message: "Unauthorized" });
  }
  const includeArchived = String(e.parameter.includeArchived || "") === "1";
  return jsonOutput_({ status: "success", forms: listAllForms_(includeArchived) });
}

// action=addForm — payload: {name}. Creates a brand-new, fully independent
// form: its own id, its own settings (all defaults until you change them),
// and its own first sheet/cycle. Nothing about any OTHER form is touched.
function handleAddForm_(payload) {
  const name = String(payload.name || "").trim();
  if (!name) return jsonOutput_({ status: "error", message: "لازم تدي الفورم اسم." });

  const forms = getFormsRegistry_();
  const id = "f_" + Utilities.getUuid().replace(/-/g, "").slice(0, 8);
  forms.push({ id, name, createdAt: new Date().toISOString(), archived: false });
  saveFormsRegistry_(forms);

  // Bootstrap its FIRST sheet/cycle right away — a form with a settings
  // blob but no sheet yet would break the moment anyone tries to view or
  // submit to it, so this makes it usable immediately after creation.
  const sheetBase = name; // admin can rename the sheet base name later from Settings, same as any form
  startNewCycle_(sheetBase, id);

  return jsonOutput_({ status: "success", form: { id, name, isDefault: false }, config: getRegConfig_(id) });
}

// action=renameForm — payload: {formId, name}. Only changes the DISPLAY
// name in the registry / the "formTitle" shown on that form's page — never
// touches its sheetBaseName or existing sheets/cycles, so nothing about
// where its data lives changes.
function handleRenameForm_(payload) {
  const formId = String(payload.formId || "").trim();
  const name = String(payload.name || "").trim();
  if (!name) return jsonOutput_({ status: "error", message: "لازم تدي الفورم اسم." });

  if (!formId) {
    // Renaming the default form just means changing its formTitle setting.
    PropertiesService.getScriptProperties().setProperty(propKey_("FORM_TITLE", ""), name);
    return jsonOutput_({ status: "success", form: { id: "", name, isDefault: true } });
  }
  const forms = getFormsRegistry_();
  const entry = forms.find(f => f.id === formId);
  if (!entry) return jsonOutput_({ status: "error", message: "الفورم ده مش موجود." });
  entry.name = name;
  saveFormsRegistry_(forms);
  return jsonOutput_({ status: "success", form: { id: formId, name, isDefault: false } });
}

// action=archiveForm — payload: {formId, archived}. Hides (or unhides) a
// form from the normal switcher WITHOUT touching its settings, sheets, or
// any registrant data — the public link and existing data both stay intact,
// it just stops showing up unless "اعرض المؤرشف" is checked. The default
// form ("") can never be archived — it's the fallback nothing else has.
function handleArchiveForm_(payload) {
  const formId = String(payload.formId || "").trim();
  if (!formId) return jsonOutput_({ status: "error", message: "الفورم الأساسي مينفعش يتأرشف." });

  const forms = getFormsRegistry_();
  const entry = forms.find(f => f.id === formId);
  if (!entry) return jsonOutput_({ status: "error", message: "الفورم ده مش موجود." });
  entry.archived = payload.archived !== false;
  saveFormsRegistry_(forms);
  return jsonOutput_({ status: "success", form: { id: formId, name: entry.name, archived: entry.archived } });
}

function getRegConfig_(formId) {
  const props = PropertiesService.getScriptProperties();
  const k = (base) => propKey_(base, formId);
  return {
    formId: formId || "",
    formTitle: props.getProperty(k("FORM_TITLE")) || "",
    sheetBaseName: props.getProperty(k("SHEET_BASE_NAME")) || (formId ? formId : SHEET_NAME),
    startAt: props.getProperty(k("REG_START")) || "",
    endAt: props.getProperty(k("REG_END")) || "",
    activeSheetName: props.getProperty(k("ACTIVE_SHEET_NAME")) || (formId ? formId : SHEET_NAME),
    cycleNumber: Number(props.getProperty(k("CYCLE_NUMBER")) || "0"),
    logoUrl: (() => {
      const fileId = props.getProperty(k("LOGO_FILE_ID"));
      return fileId ? logoUrlFromFileId_(fileId) : "";
    })(),
    sendCertAuto: props.getProperty(k("CERT_AUTO_SEND")) === "true",
    certTemplateReady: !!props.getProperty(k("CERT_TEMPLATE_FILE_ID")),
    certTemplateName: props.getProperty(k("CERT_TEMPLATE_NAME")) || "",
    fieldConfig: getFieldConfig_(formId),
    fieldDefs: buildFieldDefsForClient_(),
    fieldSections: FIELD_SECTIONS,
    customFields: getCustomFields_(formId),
    successActions: getSuccessActions_(formId),
    membershipPrefix: props.getProperty(k("MEMBERSHIP_PREFIX")) || MEMBERSHIP_PREFIX,
    // Custom confirmation email — falls back to the built-in generic text
    // (see sendConfirmationEmail_) when a form never set its own.
    confirmEmailSubject: props.getProperty(k("CONFIRM_EMAIL_SUBJECT")) || "",
    confirmEmailBody: props.getProperty(k("CONFIRM_EMAIL_BODY")) || "",
    archived: !!(getFormsRegistry_().find(f => f.id === formId) || {}).archived,
  };
}

// The original 9 toggleable fields already have their own hand-built HTML on
// the form (radio pills, national-ID auto-detect, etc.) so they're NOT part
// of this — this only describes the newer EXTRA_FIELDS, which the form
// renders generically from this metadata (see buildDynamicStep_ in
// dys_form.html). Sent to both the dashboard (to build the checkboxes) and
// the public form (to build the extra-fields step).
function buildFieldDefsForClient_() {
  const defs = {};
  Object.keys(EXTRA_FIELDS).forEach(key => {
    const f = EXTRA_FIELDS[key];
    defs[key] = { section: f.section, label: f.label, type: f.type, options: f.options || null };
  });
  defs.photo = { section: PHOTO_FIELD.section, label: PHOTO_FIELD.label, type: PHOTO_FIELD.type, options: null };
  defs.video = { section: VIDEO_FIELD.section, label: VIDEO_FIELD.label, type: VIDEO_FIELD.type, options: null };
  return defs;
}

// "before"  → now is earlier than startAt (registration hasn't opened yet)
// "closed"  → now is later than endAt (registration window is over)
// "open"    → anything else, including when no dates are configured at all
//             (keeps old deployments working exactly as before by default)
function getRegPhase_(cfg) {
  const now = Date.now();
  const start = cfg.startAt ? Date.parse(cfg.startAt) : NaN;
  const end = cfg.endAt ? Date.parse(cfg.endAt) : NaN;
  if (!isNaN(start) && now < start) return "before";
  if (!isNaN(end) && now > end) return "closed";
  return "open";
}

// Dispatches the password-protected admin actions sent via doPost. Each
// action requires ONE specific permission (not a blanket "owner" gate any
// more) — an account only needs to be granted the capability that matches
// what it's trying to do. See PERMISSION_KEYS at the top of the file for
// the full list of grantable capabilities.
// ---------------------------------------------------------------------------
// Activity log — a lightweight audit trail of every admin action, who did
// it, and when. Lives in its own "Activity Log" sheet tab (separate from
// the registration-data cycles), created automatically on first use.
// Logging failures are swallowed — an audit trail must never be the reason
// a real action fails.
// ---------------------------------------------------------------------------

const ACTIVITY_LOG_SHEET_NAME = "Activity Log";
const ACTIVITY_LOG_MAX_ROWS = 500; // trims oldest entries past this so the sheet never grows unbounded

function logActivity_(accountName, action, details) {
  try {
    const ss = getSpreadsheet_();
    let log = ss.getSheetByName(ACTIVITY_LOG_SHEET_NAME);
    if (!log) {
      log = ss.insertSheet(ACTIVITY_LOG_SHEET_NAME);
      log.appendRow(["Timestamp", "Account", "Action", "Details"]);
      log.setFrozenRows(1);
    }
    log.appendRow([new Date(), accountName || "?", action || "", details || ""]);

    const lastRow = log.getLastRow();
    if (lastRow > ACTIVITY_LOG_MAX_ROWS + 1) {
      log.deleteRows(2, lastRow - ACTIVITY_LOG_MAX_ROWS - 1);
    }
  } catch (err) {
    // logging must never break the actual action
  }
}

// Human-readable one-line summary per action, shown in the "📜 سجل النشاط"
// dashboard card — kept separate from the raw payload so nothing sensitive
// (like a newly-set password) ever ends up in the log.
function describeActionForLog_(payload) {
  switch (payload.action) {
    case "saveConfig": return "عدّل إعدادات الاستمارة العامة";
    case "uploadLogo": return "رفع شعار جديد";
    case "removeLogo": return "شال الشعار";
    case "uploadCertTemplate": return `رفع تيمبلت شهادة: ${payload.fileName || ""}`;
    case "removeCertTemplate": return "شال تيمبلت الشهادة";
    case "sendCertificate": return `بعت شهادة لسجل واحد (${payload.membershipNo || payload.rowIndex || ""})`;
    case "sendCertificatesBulk": return "بعت شهادات لدفعة من الأعضاء";
    case "sendTestCertificate": return `بعت شهادة تجريبية لـ ${payload.testEmail || ""}`;
    case "saveFieldConfig": return "عدّل إعدادات حقول الاستمارة";
    case "listAdminAccounts": return "شاف قائمة الحسابات";
    case "addAdminAccount": return `أضاف/عدّل حساب: ${payload.name || ""}`;
    case "removeAdminAccount": return `حذف حساب: ${payload.name || ""}`;
    case "reviewAccess": return payload.approve ? `وافق على دخول: ${payload.email || ""}` : `رفض دخول: ${payload.email || ""}`;
    case "updateAccountPermissions": return `عدّل صلاحيات: ${payload.name || ""}`;
    case "exportExcel": return "صدّر البيانات لملف Excel";
    case "addForm": return `أضاف فورم جديد: ${payload.name || ""}`;
    case "renameForm": return `غيّر اسم فورم لـ: ${payload.name || ""}`;
    case "archiveForm": return payload.archived === false ? "ألغى أرشفة فورم" : "أرشف فورم";
    default: return payload.action || "";
  }
}

// action=getActivityLog — returns the most recent N entries, newest first.
function handleGetActivityLog_() {
  const ss = getSpreadsheet_();
  const log = ss.getSheetByName(ACTIVITY_LOG_SHEET_NAME);
  if (!log || log.getLastRow() < 2) return jsonOutput_({ status: "success", entries: [] });
  const data = log.getRange(2, 1, log.getLastRow() - 1, 4).getValues();
  const entries = data.map(row => ({
    timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || ""),
    account: row[1],
    action: row[2],
    details: row[3],
  })).reverse();
  return jsonOutput_({ status: "success", entries });
}

// action=exportExcel — payload: {sheetName}. Formats the sheet's header row
// (bold, colored, frozen) and column widths, then exports it as a real
// .xlsx file (returned as base64 for the dashboard to trigger a download —
// Apps Script web apps can't just hand back a raw file download).
// Requires the "script.external_request" OAuth scope (same one QR codes
// need) since exporting goes through a UrlFetchApp call — see
// BACKEND_SETUP_STEPS.md if this throws an authorization error.
function handleExportExcel_(payload) {
  try {
    const sheetName = String(payload.sheetName || "").trim() || getActiveSheetName_(String(payload.formId || "").trim());
    const sheet = findSheet_(sheetName);
    if (!sheet) return jsonOutput_({ status: "error", message: `الشيت "${sheetName}" مش موجود.` });

    formatSheetForExport_(sheet);

    const ssId = getSpreadsheet_().getId();
    const gid = sheet.getSheetId();
    const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${gid}`;
    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      return jsonOutput_({ status: "error", message: "فشل التصدير — كود الاستجابة: " + resp.getResponseCode() });
    }
    const blob = resp.getBlob().setName(sheetName + ".xlsx");
    const fileBase64 = Utilities.base64Encode(blob.getBytes());
    return jsonOutput_({ status: "success", fileName: blob.getName(), fileBase64 });
  } catch (err) {
    return jsonOutput_({
      status: "error",
      message: "فشل تصدير Excel: " + String(err) +
        " — تأكد إنك مضيف صلاحية script.external_request في appsscript.json (شوف BACKEND_SETUP_STEPS.md).",
    });
  }
}

// Bold white-on-green header row + frozen row + auto-sized columns — a
// modest but real formatting pass so the exported file looks intentional
// instead of like a raw data dump.
function formatSheetForExport_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#1F6B3A")
    .setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, lastCol);
}


const ACTION_PERMISSIONS = {
  saveConfig: "manageSettings",
  uploadLogo: "manageSettings",
  removeLogo: "manageSettings",
  uploadCertTemplate: "manageCertificates",
  removeCertTemplate: "manageCertificates",
  sendCertificate: "manageCertificates",
  sendCertificatesBulk: "manageCertificates",
  sendTestCertificate: "manageCertificates",
  saveFieldConfig: "manageFields",
  listAdminAccounts: "manageAccounts",
  addAdminAccount: "manageAccounts",
  removeAdminAccount: "manageAccounts",
  updateAccountPermissions: "manageAccounts",
  reviewAccess: "manageAccounts",
  exportExcel: "manageSettings",
  getActivityLog: "manageSettings",
  addForm: "manageSettings",
  renameForm: "manageSettings",
  archiveForm: "manageSettings",
};

function handleAdminAction_(payload) {
  const requiredPermission = ACTION_PERMISSIONS[payload.action];
  if (!requiredPermission) return jsonOutput_({ status: "error", message: "Unknown admin action" });

  const account = requirePermission_(payload.password || "", requiredPermission);
  if (!account) {
    return jsonOutput_({ status: "error", message: "Unauthorized" });
  }
  if (payload.action === "saveConfig") return logAndReturn_(account, payload, handleSaveConfig_(payload));
  if (payload.action === "uploadLogo") return logAndReturn_(account, payload, handleUploadLogo_(payload));
  if (payload.action === "removeLogo") return logAndReturn_(account, payload, handleRemoveLogo_(payload));
  if (payload.action === "uploadCertTemplate") return logAndReturn_(account, payload, handleUploadCertTemplate_(payload));
  if (payload.action === "removeCertTemplate") return logAndReturn_(account, payload, handleRemoveCertTemplate_(payload));
  if (payload.action === "sendCertificate") return logAndReturn_(account, payload, handleSendCertificate_(payload));
  if (payload.action === "sendCertificatesBulk") return logAndReturn_(account, payload, handleSendCertificatesBulk_(payload));
  if (payload.action === "sendTestCertificate") return logAndReturn_(account, payload, handleSendTestCertificate_(payload));
  if (payload.action === "saveFieldConfig") return logAndReturn_(account, payload, handleSaveFieldConfig_(payload));
  if (payload.action === "listAdminAccounts") return handleListAdminAccounts_(); // read-only, not logged — keeps the log focused on actual changes
  if (payload.action === "addAdminAccount") return logAndReturn_(account, payload, handleAddAdminAccount_(payload));
  if (payload.action === "removeAdminAccount") return logAndReturn_(account, payload, handleRemoveAdminAccount_(payload));
  if (payload.action === "reviewAccess") return logAndReturn_(account, payload, handleDashboardReviewAccess_(payload));
  if (payload.action === "updateAccountPermissions") return logAndReturn_(account, payload, handleUpdateAccountPermissions_(payload));
  if (payload.action === "exportExcel") return logAndReturn_(account, payload, handleExportExcel_(payload));
  if (payload.action === "getActivityLog") return handleGetActivityLog_(); // read-only, not logged
  if (payload.action === "addForm") return logAndReturn_(account, payload, handleAddForm_(payload));
  if (payload.action === "renameForm") return logAndReturn_(account, payload, handleRenameForm_(payload));
  if (payload.action === "archiveForm") return logAndReturn_(account, payload, handleArchiveForm_(payload));
  return jsonOutput_({ status: "error", message: "Unknown admin action" });
}

// Logs the action (skips logging if the handler itself reported an error —
// failed attempts aren't useful audit history the way successful changes
// are) and passes the handler's own response straight through unchanged.
function logAndReturn_(account, payload, response) {
  try {
    const parsed = JSON.parse(response.getContent());
    if (parsed.status === "success") {
      logActivity_(account.name, payload.action, describeActionForLog_(payload));
    }
  } catch (e) { /* if we can't tell whether it succeeded, don't log a guess */ }
  return response;
}

// Strips passwords for anything sent back to the dashboard.
function publicAccount_(a) {
  return { name: a.name, role: a.role, permissions: a.permissions, email: a.email || "", status: a.status || "approved" };
}

// action=listAdminAccounts — names + permissions only, NEVER passwords.
function handleListAdminAccounts_() {
  const accounts = getAdminAccounts_().map(publicAccount_);
  return jsonOutput_({ status: "success", accounts });
}

// action=addAdminAccount — payload: {name, password, permissions: {manageSettings,
// manageFields, manageCertificates, manageAccounts}}. Adding an account with
// a name that already exists overwrites that account (lets you change
// someone's password/permissions without a separate "edit" action).
function handleAddAdminAccount_(payload) {
  const name = String(payload.name || "").trim();
  const password = String(payload.newPassword || "").trim();
  const permissions = sanitizePermissions_(payload.permissions);
  if (!name || !password) {
    return jsonOutput_({ status: "error", message: "لازم اسم وكلمة سر." });
  }
  if (password.length < 4) {
    return jsonOutput_({ status: "error", message: "كلمة السر قصيرة أوي — اكتب حاجة أطول." });
  }
  const accounts = getAdminAccounts_().filter(a => a.name !== name);
  accounts.push({ name, password, permissions });
  saveAdminAccounts_(accounts);
  return jsonOutput_({ status: "success", accounts: accounts.map(publicAccount_) });
}

// action=removeAdminAccount — payload: {name}. Refuses to remove the last
// remaining account that can manage accounts, so you can never lock
// everyone out of the "👥 حسابات الدخول" card entirely.
function handleRemoveAdminAccount_(payload) {
  const name = String(payload.name || "").trim();
  const accounts = getAdminAccounts_();
  const target = accounts.find(a => a.name === name);
  if (!target) return jsonOutput_({ status: "error", message: "الحساب ده مش موجود." });

  const managerCount = accounts.filter(a => a.permissions.manageAccounts).length;
  if (target.permissions.manageAccounts && managerCount <= 1) {
    return jsonOutput_({ status: "error", message: "ده آخر حساب معاه صلاحية إدارة الحسابات — مينفعش تمسحه عشان متتقفلش برة النظام." });
  }

  const remaining = accounts.filter(a => a.name !== name);
  saveAdminAccounts_(remaining);
  return jsonOutput_({ status: "success", accounts: remaining.map(publicAccount_) });
}

// action=reviewAccess (POST, dashboard-side) — payload: {email, approve}.
// Same effect as clicking the "قبول"/"رفض" link in the notification email,
// just from inside "👥 حسابات الدخول" for whoever's already logged in with
// manageAccounts — handy if the email never arrives or they're already in
// the dashboard when a request comes in.
function handleDashboardReviewAccess_(payload) {
  const email = String(payload.email || "").toLowerCase().trim();
  const accounts = getRawAdminAccounts_();
  const idx = accounts.findIndex(a => a.email === email && a.status === "pending");
  if (idx === -1) return jsonOutput_({ status: "error", message: "الطلب ده مش موجود (يمكن اتراجع أو اتوافق عليه خلاص)." });

  if (payload.approve) {
    delete accounts[idx].status;
    delete accounts[idx].approvalToken;
  } else {
    accounts.splice(idx, 1);
  }
  saveAdminAccounts_(accounts);
  return jsonOutput_({ status: "success", accounts: getAdminAccounts_().map(publicAccount_) });
}

// action=updateAccountPermissions — payload: {name, permissions}. Changes
// an EXISTING account's permissions in place — unlike handleAddAdminAccount_
// (password accounts only, requires a new password every time), this works
// for ANY account, password-based or Google-signed-in, and never touches
// its password/email/other identity fields. This is what the dashboard's
// inline permission checkboxes next to each account call — no more
// delete-and-re-add dance to change what a Google account can do.
function handleUpdateAccountPermissions_(payload) {
  const name = String(payload.name || "").trim();
  const accounts = getRawAdminAccounts_();
  const idx = accounts.findIndex(a => a.name === name);
  if (idx === -1) {
    return jsonOutput_({
      status: "error",
      message: name === "الحساب الأساسي"
        ? "\"الحساب الأساسي\" (كلمة السر القديمة) دايمًا معاه كل الصلاحيات ومينفعش يتعدّل."
        : "الحساب ده مش موجود.",
    });
  }
  if (accounts[idx].status === "pending") {
    return jsonOutput_({ status: "error", message: "الحساب ده لسه بانتظار الموافقة — وافق عليه الأول." });
  }

  const current = normalizeAccount_(accounts[idx]).permissions;
  const newPermissions = sanitizePermissions_(payload.permissions);
  const managerCountWithoutThis = getAdminAccounts_().filter(a => a.name !== name && a.permissions.manageAccounts).length;
  if (current.manageAccounts && !newPermissions.manageAccounts && managerCountWithoutThis === 0) {
    return jsonOutput_({ status: "error", message: "ده آخر حساب معاه صلاحية إدارة الحسابات — مينفعش تشيلها عشان متتقفلش برة النظام." });
  }

  accounts[idx].permissions = newPermissions;
  saveAdminAccounts_(accounts);
  return jsonOutput_({ status: "success", accounts: getAdminAccounts_().map(publicAccount_) });
}

// Forces an arbitrary incoming permissions object into exactly the shape we
// trust: every key in PERMISSION_KEYS explicitly true/false, nothing else.
function sanitizePermissions_(incoming) {
  const src = incoming || {};
  const out = {};
  PERMISSION_KEYS.forEach(key => { out[key] = src[key] === true; });
  return out;
}

// Saves form title / sheet base name / registration window, and — only if
// asked to (startNewCycle: true), or if there's no active sheet at all yet —
// creates a brand-new sheet tab and switches the form to write into it.
// This is exactly the "open the form again → it registers into a new Google
// Sheet automatically" behaviour: each cycle gets its own tab, nothing from
// a previous cycle is ever touched or overwritten.
function handleSaveConfig_(payload) {
  const props = PropertiesService.getScriptProperties();
  const formId = String(payload.formId || "").trim();
  const k = (base) => propKey_(base, formId);

  const formTitle = String(payload.formTitle || "").trim();
  const sheetBaseName = String(payload.sheetBaseName || "").trim();
  const startAt = String(payload.startAt || "").trim();
  const endAt = String(payload.endAt || "").trim();

  props.setProperty(k("FORM_TITLE"), formTitle);
  if (sheetBaseName) props.setProperty(k("SHEET_BASE_NAME"), sheetBaseName);
  props.setProperty(k("REG_START"), startAt);
  props.setProperty(k("REG_END"), endAt);
  // "sendCertAuto" only arrives when the dashboard sends it (see collectConfigPayload
  // in dys_dashboard.html) — if payload doesn't include the key at all we leave the
  // stored value untouched, but the dashboard always sends it explicitly, so this
  // simply mirrors whatever the toggle in Settings was set to.
  if (typeof payload.sendCertAuto !== "undefined") {
    props.setProperty(k("CERT_AUTO_SEND"), payload.sendCertAuto ? "true" : "false");
  }

  // Membership number prefix ("بادئة رقم العضوية") — letters/digits only,
  // 2–8 chars, uppercased. Falls back to the global default (MEMBERSHIP_PREFIX)
  // if left blank or sent invalid, so a bad value never breaks numbering.
  if (typeof payload.membershipPrefix !== "undefined") {
    const cleanedPrefix = String(payload.membershipPrefix || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
    props.setProperty(k("MEMBERSHIP_PREFIX"), (cleanedPrefix.length >= 2 && cleanedPrefix.length <= 8) ? cleanedPrefix : "");
  }

  // Custom confirmation email text — blank means "use the built-in generic
  // text" (see sendConfirmationEmail_), so an empty string is a valid,
  // intentional value here, not something to skip.
  if (typeof payload.confirmEmailSubject !== "undefined") {
    props.setProperty(k("CONFIRM_EMAIL_SUBJECT"), String(payload.confirmEmailSubject || "").trim());
  }
  if (typeof payload.confirmEmailBody !== "undefined") {
    props.setProperty(k("CONFIRM_EMAIL_BODY"), String(payload.confirmEmailBody || "").trim());
  }

  // Success-screen buttons ("🎉 خيارات بعد التسجيل"). Only touched when the
  // dashboard actually sends the key (it always does from the Settings tab
  // save button — see collectConfigPayload in dys_dashboard.html) — so a
  // partial/older client calling saveConfig for something else never wipes
  // this list out.
  if (Array.isArray(payload.successActions)) {
    const cleaned = payload.successActions
      .map(sanitizeSuccessAction_)
      .filter(Boolean);
    saveSuccessActions_(cleaned, formId);
  }

  let activeSheetName = props.getProperty(k("ACTIVE_SHEET_NAME")) || "";

  if (payload.startNewCycle || !activeSheetName) {
    activeSheetName = startNewCycle_(sheetBaseName || props.getProperty(k("SHEET_BASE_NAME")) || (formId || SHEET_NAME), formId);
  }

  return jsonOutput_({ status: "success", config: getRegConfig_(formId) });
}

// Creates a new sheet tab named after the base name (first cycle keeps the
// base name as-is, later cycles get " 2", " 3", ... appended so the name
// stays readable), sets it as the ACTIVE sheet, and bumps CYCLE_NUMBER — all
// scoped to the given formId (see propKey_ above).
function startNewCycle_(base, formId) {
  const props = PropertiesService.getScriptProperties();
  const ss = getSpreadsheet_();
  const k = (b) => propKey_(b, formId);

  let cycle = Number(props.getProperty(k("CYCLE_NUMBER")) || "0") + 1;
  let candidate = cycle === 1 ? base : `${base} ${cycle}`;
  while (ss.getSheetByName(candidate)) {
    cycle += 1;
    candidate = `${base} ${cycle}`;
  }

  const sheet = ss.insertSheet(candidate);
  sheet.appendRow(HEADERS);
  sheet.setFrozenRows(1);

  props.setProperty(k("CYCLE_NUMBER"), String(cycle));
  props.setProperty(k("ACTIVE_SHEET_NAME"), candidate);
  props.setProperty(k("SHEET_BASE_NAME"), base);
  addFormSheetName_(formId, candidate);

  return candidate;
}

// Decodes a base64 image sent from the dashboard's Settings tab, stores it
// in Google Drive (shared "anyone with the link can view" so the public
// form can display it), and remembers its URL. The previous logo file (if
// any) is trashed so you don't end up with a pile of old images in Drive.
function handleUploadLogo_(payload) {
  try {
    const formId = String(payload.formId || "").trim();
    const raw = String(payload.imageBase64 || "");
    const commaIdx = raw.indexOf(",");
    const base64 = commaIdx > -1 && raw.slice(0, commaIdx).indexOf("base64") > -1 ? raw.slice(commaIdx + 1) : raw;
    const mimeType = payload.mimeType || "image/png";
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mimeType, payload.fileName || "dys-form-logo");

    const props = PropertiesService.getScriptProperties();
    const k = (b) => propKey_(b, formId);
    const oldFileId = props.getProperty(k("LOGO_FILE_ID"));
    if (oldFileId) {
      try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e2) { /* already gone — fine */ }
    }

    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId = file.getId();

    props.setProperty(k("LOGO_FILE_ID"), fileId);
    props.deleteProperty(k("LOGO_URL")); // legacy key — URL is now always derived from LOGO_FILE_ID, see logoUrlFromFileId_()

    return jsonOutput_({ status: "success", logoUrl: logoUrlFromFileId_(fileId) });
  } catch (err) {
    return jsonOutput_({ status: "error", message: String(err) });
  }
}

// `drive.google.com/uc?export=view&id=...` (the old format this used to
// generate) is Google's classic direct-file link, but it's unreliable for
// hotlinking as an <img src> — it can silently break, get rate-limited, or
// redirect to a "can't scan for viruses" interstitial instead of the image.
// `drive.google.com/thumbnail?id=...` is the endpoint Drive itself uses for
// image previews and is far more reliable for this use case.
function logoUrlFromFileId_(fileId) {
  return `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000`;
}

// Removes the custom logo — the form falls back to its built-in badge.
function handleRemoveLogo_(payload) {
  const formId = String((payload && payload.formId) || "").trim();
  const props = PropertiesService.getScriptProperties();
  const k = (b) => propKey_(b, formId);
  const oldFileId = props.getProperty(k("LOGO_FILE_ID"));
  if (oldFileId) {
    try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e2) { /* already gone — fine */ }
  }
  props.deleteProperty(k("LOGO_FILE_ID"));
  props.deleteProperty(k("LOGO_URL"));
  return jsonOutput_({ status: "success" });
}


// ---------------------------------------------------------------------------
// Certificates (PDF, generated from an uploaded Word template on each send)
// ---------------------------------------------------------------------------
// How it works:
//   1. You upload a .docx certificate template from the dashboard's Settings
//      tab. It gets converted to a Google Doc and stored in Drive — this
//      conversion is what lets us fill it in and export a PDF later.
//   2. Anywhere in the template text, you type placeholders like {{name}}
//      or {{membershipNo}} — see PLACEHOLDER FIELDS below for the full list.
//      They can be anywhere: inside a text box, a table cell, styled with
//      any font/color you want — replaceText() finds them regardless.
//   3. When a certificate needs to be sent, we duplicate the template,
//      swap in the real values, export that copy as a PDF, email it, then
//      delete the temporary copy (your original template is never touched).
//
// PLACEHOLDER FIELDS you can use inside the template — matches every column
// in the registration sheet, so you decide which ones actually appear on the
// certificate and where:
//   {{name}}             — الاسم
//   {{membershipNo}}      — رقم العضوية (OSH-000123)
//   {{age}}               — السن
//   {{gender}}            — النوع (Male/Female)
//   {{nationalId}}        — الرقم القومي
//   {{phone}}             — رقم الهاتف
//   {{whatsapp}}          — رقم الواتساب
//   {{email}}             — الإيميل
//   {{faculty}}           — الكلية
//   {{graduate}}          — خريج ولا لأ (Yes/No)
//   {{role}}              — الدور داخل الكيان
//   {{hasJob}}             — بيشتغل ولا لأ (Yes/No)
//   {{currentJob}}         — الوظيفة الحالية
//   {{registrationDate}}   — تاريخ التسجيل الأصلي (من عمود Timestamp)
//   {{date}}               — تاريخ إرسال الشهادة (بيتحسب لحظة الإرسال، مش وقت التسجيل)
//
//   PLUS: أي حقل من الحقول الإضافية اللي فعّلتها من "🧩 حقول الاستمارة" — مثلاً
//   {{birthDate}}, {{governorate}}, {{tshirtSize}} ... (شوف EXTRA_FIELDS فوق
//   للقائمة الكاملة)، وأي حقل مخصص ضفته بنفسك كـ {{c_xxxxx}} — الكود الدقيق
//   بيتعرض جنب كل حقل مخصص في نفس الكارت لما تضيفه.
//
// ⚠️ Requires the "Drive API" advanced service to be enabled once in this
// project (Apps Script editor ▸ Services ▸ + ▸ Drive API). This is what
// lets a .docx upload be converted into an editable Google Doc — see
// BACKEND_SETUP_STEPS.md for the exact steps.

function handleUploadCertTemplate_(payload) {
  try {
    const formId = String(payload.formId || "").trim();
    const raw = String(payload.fileBase64 || "");
    const commaIdx = raw.indexOf(",");
    const base64 = commaIdx > -1 && raw.slice(0, commaIdx).indexOf("base64") > -1 ? raw.slice(commaIdx + 1) : raw;
    const mimeType = payload.mimeType || "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const fileName = payload.fileName || "dys-certificate-template.docx";
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);

    const props = PropertiesService.getScriptProperties();
    const k = (b) => propKey_(b, formId);
    const oldFileId = props.getProperty(k("CERT_TEMPLATE_FILE_ID"));

    // Convert the uploaded .docx straight into a Google Doc so we can fill
    // it in with replaceText() later. Needs the Drive API advanced service.
    //
    // NOTE: enabling "Drive API" in Apps Script today binds to Drive API v3,
    // not the old v2. In v3, Files.insert() was renamed Files.create(), it
    // takes (requestBody, media) instead of (resource, blob, optionalArgs),
    // "title" became "name", and there's no separate {convert:true} flag —
    // conversion happens automatically because requestBody.mimeType asks for
    // a native Google Workspace format while the uploaded media is .docx.
    const converted = Drive.Files.create(
      { name: fileName.replace(/\.docx$/i, ""), mimeType: MimeType.GOOGLE_DOCS },
      blob
    );

    if (oldFileId) {
      try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e2) { /* already gone — fine */ }
    }

    props.setProperty(k("CERT_TEMPLATE_FILE_ID"), converted.id);
    props.setProperty(k("CERT_TEMPLATE_NAME"), fileName);

    return jsonOutput_({ status: "success", fileName });
  } catch (err) {
    return jsonOutput_({
      status: "error",
      message: "فشل رفع التيمبلت: " + String(err) +
        " — تأكد إنك فعّلت Drive API من Services جوه محرر الكود (شوف BACKEND_SETUP_STEPS.md).",
    });
  }
}

function handleRemoveCertTemplate_(payload) {
  const formId = String((payload && payload.formId) || "").trim();
  const props = PropertiesService.getScriptProperties();
  const k = (b) => propKey_(b, formId);
  const oldFileId = props.getProperty(k("CERT_TEMPLATE_FILE_ID"));
  if (oldFileId) {
    try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e2) { /* already gone — fine */ }
  }
  props.deleteProperty(k("CERT_TEMPLATE_FILE_ID"));
  props.deleteProperty(k("CERT_TEMPLATE_NAME"));
  props.setProperty(k("CERT_AUTO_SEND"), "false");
  return jsonOutput_({ status: "success" });
}

// Fills the template with the given field values and returns a PDF blob.
// `fields` keys must match the {{placeholder}} names used in the template.
function generateCertificatePdf_(fields, formId) {
  const templateId = PropertiesService.getScriptProperties().getProperty(propKey_("CERT_TEMPLATE_FILE_ID", formId));
  if (!templateId) throw new Error("لسه مفيش تيمبلت شهادة متحمل من تبويب الإعدادات.");

  const templateFile = DriveApp.getFileById(templateId);
  if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
    // The upload step asked Drive to convert the .docx into a Google Doc —
    // if this ever isn't true, the conversion silently didn't happen and
    // DocumentApp.openById() below would fail with a much more confusing
    // error, so catch it here with a clear message instead.
    throw new Error("التيمبلت المحفوظ مش Google Doc (mimeType: " + templateFile.getMimeType() +
      ") — جرب تمسحه وترفعه تاني من تبويب الإعدادات.");
  }

  const copy = templateFile.makeCopy("Certificate - temp - " + new Date().getTime());
  try {
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    Object.keys(fields).forEach(key => {
      body.replaceText("\\{\\{" + key + "\\}\\}", escapeForReplaceText_(fields[key]));
    });

    // {{qrcode}} and {{photo}} are IMAGE placeholders, not text — handled
    // separately after the text substitutions above. A template only needs
    // to contain the token if you actually want that image on it; templates
    // without the token are completely unaffected (findText just finds
    // nothing and insertImagePlaceholder_ silently does nothing).
    if (fields.qrPayload) {
      const qrBlob = fetchQrCodeBlob_(fields.qrPayload);
      insertImagePlaceholder_(body, "{{qrcode}}", qrBlob, 90, 90);
    }
    if (fields.photoUrl) {
      const photoBlob = fetchImageBlobFromUrl_(fields.photoUrl);
      insertImagePlaceholder_(body, "{{photo}}", photoBlob, 110, 130);
    }

    doc.saveAndClose();

    const pdfBlob = DriveApp.getFileById(copy.getId()).getAs(MimeType.PDF);
    pdfBlob.setName((fields.name || "Certificate") + ".pdf");
    return pdfBlob;
  } finally {
    try { DriveApp.getFileById(copy.getId()).setTrashed(true); } catch (e2) { /* best-effort cleanup */ }
  }
}

// Replaces a {{token}} in the doc body with an inline image, if both the
// token and the image are present — a template with no {{qrcode}}/{{photo}}
// token, or a registrant with no photo, just silently does nothing here.
// Requires the "https://www.googleapis.com/auth/script.external_request"
// OAuth scope for fetchQrCodeBlob_/fetchImageBlobFromUrl_ to work — see
// BACKEND_SETUP_STEPS.md if this throws an authorization error the same way
// the {{documents}} scope did earlier.
function insertImagePlaceholder_(body, token, blob, widthPt, heightPt) {
  if (!blob) return;
  const found = body.findText(token);
  if (!found) return; // template doesn't use this placeholder — nothing to do
  const textEl = found.getElement().asText();
  const startOffset = found.getStartOffset();
  const endOffsetInclusive = found.getEndOffsetInclusive();
  textEl.deleteText(startOffset, endOffsetInclusive);
  const image = textEl.insertInlineImage(startOffset, blob);
  if (widthPt) image.setWidth(widthPt);
  if (heightPt) image.setHeight(heightPt);
}

// Free, no-API-key QR generator (api.qrserver.com). `data` is whatever text
// should be encoded — we encode the registrant's name + membership number +
// issue date directly INTO the QR code itself (no lookup server needed), so
// scanning it shows the certified info even without a live verification
// page. Returns null on any network failure (never blocks certificate
// generation over a QR image failing).
function fetchQrCodeBlob_(data) {
  try {
    const url = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(data);
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    return resp.getBlob();
  } catch (err) {
    return null;
  }
}

function fetchImageBlobFromUrl_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    return resp.getBlob();
  } catch (err) {
    return null;
  }
}

// replaceText() treats its first argument as a regex — escape anything a
// real name/value might contain (parentheses, dots, etc.) so it's matched
// as literal text instead.
function escapeForReplaceText_(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Builds the {{placeholder}} field map + sends one certificate email.
// Lets errors PROPAGATE (doesn't catch) — callers that want a real error
// message (the test-send button, single-send) call this directly; callers
// doing a best-effort batch (bulk send, auto-send-on-registration) go
// through sendCertificateEmail_() below instead, which never throws.
//
// `p` can be either a raw form submission payload (camelCase keys, used
// right after a live registration) or a person object built from a sheet
// row via rowToPerson_() (used for manual/bulk sends) — both shapes carry
// the same field names, so this works for either.
function sendCertificateEmailCore_(p, membershipNo, formId) {
  const email = String(p.email || "").trim();
  if (!email) throw new Error("السجل ده مفيهوش إيميل.");
  const str = (v) => String(v || "").trim();

  const fields = {
    name: str(p.name),
    membershipNo: str(membershipNo),
    age: str(p.age),
    gender: str(p.gender),
    nationalId: str(p.nationalId),
    phone: str(p.phone),
    whatsapp: str(p.whatsapp),
    email: email,
    faculty: str(p.faculty),
    graduate: str(p.graduate),
    role: str(p.committee),
    hasJob: str(p.hasJob),
    currentJob: str(p.currentJob),
    // registrationDate: when they actually registered (from the sheet's
    // Timestamp column) — falls back to "now" for a fresh live submission,
    // since there's no Timestamp column value yet at that point.
    registrationDate: p.registrationDate
      ? str(p.registrationDate)
      : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
    // date: when the CERTIFICATE is being sent — always "now".
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
  };

  // Every EXTRA_FIELDS key (address, birthDate, tshirtSize, ...) becomes a
  // usable {{key}} placeholder automatically — no need to list them by hand
  // here the way the original fields above are.
  Object.keys(EXTRA_FIELDS).forEach(key => { fields[key] = str(p[key]); });

  // Custom fields become {{c_xxxxx}} placeholders — see sanitizeCustomField_.
  const customValues = p.customFields || {};
  getCustomFields_(formId).forEach(cf => { fields[cf.key] = str(customValues[cf.key]); });

  // {{qrcode}} and {{photo}} — image placeholders, handled inside
  // generateCertificatePdf_ (NOT regular text substitution). qrPayload is
  // the actual text encoded into the QR image, not a placeholder key itself.
  fields.qrPayload = `أسرة صناع الحياة | ${fields.name} | عضوية ${fields.membershipNo} | ${fields.date}`;
  fields.photoUrl = str(p.photoUrl);

  const pdf = generateCertificatePdf_(fields, formId);
  sendEmail_(
    email,
    "شهادتك — أسرة صناع الحياة",
    `أهلًا ${fields.name}،\n\n` +
    `تحية طيبة، مرفق شهادتك.\n\n` +
    `تحياتنا،\nفريق أسرة صناع الحياة`,
    { attachments: [pdf] }
  );
}

// Best-effort wrapper around sendCertificateEmailCore_() — swallows errors
// and returns true/false. Use this for batch/automatic sends where one
// failure shouldn't blow up the whole run and there's no one watching for
// a detailed error message in the moment.
function sendCertificateEmail_(p, membershipNo, formId) {
  try {
    sendCertificateEmailCore_(p, membershipNo, formId);
    return true;
  } catch (err) {
    console.error("Certificate email failed:", err);
    return false;
  }
}

// Maps a raw sheet row + its header row into the same field shape
// sendCertificateEmail_() expects, by header NAME (not position) — so this
// keeps working even if columns get reordered.
function rowToPerson_(headers, row) {
  const get = (headerName) => {
    const i = headers.indexOf(headerName);
    return i > -1 ? row[i] : "";
  };
  const ts = get("Timestamp");
  const person = {
    name: get("Name"),
    email: get("Email"),
    age: get("Age"),
    gender: get("Gender"),
    nationalId: get("National ID"),
    phone: get("Phone"),
    whatsapp: get("Whatsapp"),
    address: get("Address"),
    faculty: get("Faculty"),
    graduate: get("Graduate"),
    hasJob: get("Has Job"),
    currentJob: get("Current Job"),
    registrationDate: ts instanceof Date
      ? Utilities.formatDate(ts, Session.getScriptTimeZone(), "dd/MM/yyyy")
      : String(ts || ""),
    committee: get("Role in Entity"),
    membershipNo: get("Membership No"),
  };
  // Generic EXTRA_HEADERS fields, e.g. person.birthDate, person.tshirtSize, ...
  EXTRA_HEADERS.forEach(h => { person[HEADER_TO_EXTRA_FIELD[h]] = get(h); });
  // Custom fields, available to certificate templates as {{c_xxxxx}} — see
  // sendCertificateEmailCore_'s placeholder resolution.
  try {
    person.customFields = JSON.parse(get(CUSTOM_FIELDS_HEADER) || "{}");
  } catch (e) {
    person.customFields = {};
  }
  person.photoUrl = get("Photo URL");
  return person;
}

// action=sendCertificate — single certificate, looked up by National ID
// inside the given (or currently active) sheet/cycle.
function handleSendCertificate_(payload) {
  const formId = String(payload.formId || "").trim();
  const cfg = getRegConfig_(formId);
  if (!cfg.certTemplateReady) {
    return jsonOutput_({ status: "error", message: "ارفع تيمبلت الشهادة الأول من تبويب الإعدادات." });
  }
  const sheet = findSheet_(payload.sheet || cfg.activeSheetName);
  if (!sheet) return jsonOutput_({ status: "error", message: "الشيت مش موجود." });

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const nidCol = headers.indexOf("National ID");
  const nationalId = String(payload.nationalId || "").trim();
  const row = data.slice(1).find(r => String(r[nidCol]).trim() === nationalId);
  if (!row) return jsonOutput_({ status: "error", message: "السجل مش موجود في الشيت ده." });

  const person = rowToPerson_(headers, row);
  if (!person.email) return jsonOutput_({ status: "error", message: "السجل ده مفيهوش إيميل." });

  try {
    sendCertificateEmailCore_(person, person.membershipNo, formId);
    return jsonOutput_({ status: "success", message: "اترسلت الشهادة ✓" });
  } catch (err) {
    return jsonOutput_({ status: "error", message: "فشل إرسال الشهادة: " + String(err) });
  }
}

// action=sendCertificatesBulk — sends to everyone with a valid email in the
// given (or currently active) sheet/cycle. Best-effort per row: one failure
// doesn't stop the rest.
function handleSendCertificatesBulk_(payload) {
  const formId = String(payload.formId || "").trim();
  const cfg = getRegConfig_(formId);
  if (!cfg.certTemplateReady) {
    return jsonOutput_({ status: "error", message: "ارفع تيمبلت الشهادة الأول من تبويب الإعدادات." });
  }
  const sheet = findSheet_(payload.sheet || cfg.activeSheetName);
  if (!sheet) return jsonOutput_({ status: "error", message: "الشيت مش موجود." });

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const rows = data.slice(1).filter(r => r.some(c => String(c).trim() !== ""));

  let sent = 0, failed = 0, skippedNoEmail = 0;
  rows.forEach(row => {
    const person = rowToPerson_(headers, row);
    if (!person.email) { skippedNoEmail += 1; return; }
    if (sendCertificateEmail_(person, person.membershipNo, formId)) sent += 1; else failed += 1;
  });

  return jsonOutput_({ status: "success", sent, failed, skippedNoEmail, total: rows.length });
}

// action=sendTestCertificate — sends one certificate with sample data to an
// email the admin types in, WITHOUT touching the sheet. Use this to check a
// newly uploaded template's layout/placeholders before trusting it with a
// real batch.
function handleSendTestCertificate_(payload) {
  const formId = String(payload.formId || "").trim();
  const cfg = getRegConfig_(formId);
  if (!cfg.certTemplateReady) {
    return jsonOutput_({ status: "error", message: "ارفع تيمبلت الشهادة الأول من تبويب الإعدادات." });
  }
  const email = String(payload.email || "").trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return jsonOutput_({ status: "error", message: "اكتب إيميل صحيح للاختبار." });
  }
  const sample = {
    name: "اسم تجريبي",
    email: email,
    faculty: "كلية تجريبية",
    committee: "لجنة تجريبية",
    nationalId: "00000000000000",
  };
  try {
    sendCertificateEmailCore_(sample, "OSH-000000", formId);
    return jsonOutput_({ status: "success", message: "اترسلت شهادة تجريبية ✓ — روح شوف إيميلك." });
  } catch (err) {
    return jsonOutput_({ status: "error", message: "فشل إرسال الشهادة التجريبية: " + String(err) });
  }
}


// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

function getActiveSheetName_(formId) {
  return PropertiesService.getScriptProperties().getProperty(propKey_("ACTIVE_SHEET_NAME", formId)) || (formId || SHEET_NAME);
}

// Looks up a sheet WITHOUT creating it. Used anywhere we must never silently
// spawn a new empty tab just because someone passed an unexpected name.
function findSheet_(name) {
  return getSpreadsheet_().getSheetByName(name);
}

// Gets (and creates, if missing) the CURRENT active sheet for the given form
// — i.e. the one THAT form's registration cycle is writing into. On a fresh
// install (formId ""), this is what bootstraps ACTIVE_SHEET_NAME/
// SHEET_BASE_NAME the very first time, so old deployments that never touch
// the new Settings tab keep behaving exactly like before.
function getSheet_(formId, nameOpt) {
  const name = nameOpt || getActiveSheetName_(formId);
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);

    const props = PropertiesService.getScriptProperties();
    const k = (b) => propKey_(b, formId);
    if (!props.getProperty(k("ACTIVE_SHEET_NAME"))) props.setProperty(k("ACTIVE_SHEET_NAME"), name);
    if (!props.getProperty(k("SHEET_BASE_NAME"))) props.setProperty(k("SHEET_BASE_NAME"), name);
    addFormSheetName_(formId, name);
  } else {
    healSheetHeaders_(sheet);
  }
  return sheet;
}

// Self-heals an EXISTING sheet/cycle whenever HEADERS grows (like it just
// did — 17 new columns for the extra fields + custom fields). Only ever
// WRITES into columns that are currently blank — if a column already has
// some other value there (a real conflict), this leaves it alone and
// checkHeaders() will report it instead, so real data is never clobbered.
function healSheetHeaders_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastCol >= HEADERS.length) return; // already has room for every header — nothing to do
  const range = sheet.getRange(1, 1, 1, HEADERS.length);
  const actual = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const full = HEADERS.map((h, i) => (actual[i] ? actual[i] : h));
  range.setValues([full]);
}

function buildRow_(p, membershipNo) {
  const str = (v) => String(v || "").trim();
  // Order MUST match HEADERS above. This first part (15 columns) is
  // unchanged from before — the newer EXTRA_HEADERS + custom-fields column
  // are appended generically right after it.
  const legacyRow = [
    new Date(),
    membershipNo,
    str(p.name),
    str(p.age),
    p.gender,
    str(p.nationalId),
    str(p.phone),
    str(p.whatsapp),
    str(p.email),
    str(p.address), // now a real configurable field again — see EXTRA_FIELDS.address
    str(p.faculty),
    p.graduate,
    str(p.committee),
    p.hasJob,
    str(p.currentJob),
  ];
  const extraRow = EXTRA_HEADERS.map(h => str(p[HEADER_TO_EXTRA_FIELD[h]]));
  const customJson = JSON.stringify(p.customFields || {});
  return legacyRow.concat(extraRow).concat([customJson]).concat([str(p.photoUrl), str(p.videoUrl)]);
}

function isDuplicateNid_(nationalId, formId) {
  const sheet = getSheet_(formId);
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const nidCol = headers.indexOf("National ID");
  if (nidCol === -1) return false;
  return data.slice(1).some(row => String(row[nidCol]).trim() === nationalId);
}

// ---------------------------------------------------------------------------
// Event check-in (QR scanner) — see dys_checkin.html. Lets anyone with a
// valid dashboard password scan a member's certificate QR code at the door
// of an event and mark them "arrived" in real time. Reuses everything that
// already exists (same sheet, same "Membership No" column, same QR payload
// the certificates already embed — no certificate re-issue needed) and only
// adds one new column, "Checked In At", which self-heals into every
// existing sheet automatically via healSheetHeaders_ the first time it's
// touched. Nothing about registration, certificates, or the dashboard table
// changes because of this — it's purely additive.

// The certificate's QR currently encodes a decorative line like
// "أسرة صناع الحياة | يوسف محروس | عضوية OSH-000456 | ..." (see
// fields.qrPayload below) — this just pulls the "OSH-000456" part out of
// whatever text the scanner read, so it works with certificates already
// issued before this feature existed, AND with someone just typing the
// membership number in manually as a fallback if the camera can't get a
// clean scan.
function extractMembershipNo_(raw) {
  const s = String(raw || "").toUpperCase();
  // Prefix-agnostic on purpose — different forms can have different
  // membership prefixes (see cfg.membershipPrefix / MEMBERSHIP_PREFIX
  // default), and check-in doesn't necessarily know in advance which form a
  // given code belongs to.
  const m = s.match(/[A-Z]+-\d+/);
  if (m) return m[0];
  // Fallback: bare digits typed by hand, e.g. "456" or "000456" — pad it out
  // to match the default prefix's format instead of failing to find a match.
  const digits = s.match(/\d+/);
  if (digits) return `${MEMBERSHIP_PREFIX}-${digits[0].padStart(6, "0")}`;
  return "";
}

function rowToMember_(headers, row) {
  const get = (name) => { const i = headers.indexOf(name); return i > -1 ? row[i] : ""; };
  return {
    membershipNo: get("Membership No"),
    name: get("Name"),
    role: get("Role in Entity"),
    faculty: get("Faculty"),
    photoUrl: get("Photo URL"),
  };
}

// formId "" (or omitted) searches every sheet across every form, same as
// before multi-form support existed. A specific formId restricts the search
// to only that form's own sheets/cycles (matched the same way
// handleListCycles_ matches them — by sheetBaseName).
function findMemberRowInAllCycles_(code, formId) {
  const ss = getSpreadsheet_();
  const cfg = getRegConfig_(formId || "");
  const activeName = cfg.activeSheetName;

  let candidateNames;
  if (formId) {
    const base = cfg.sheetBaseName;
    candidateNames = [activeName].concat(
      ss.getSheets().map(s => s.getName())
        .filter(n => n !== activeName && (n === base || n.indexOf(base + " ") === 0))
    );
  } else {
    // Check the active cycle first (the common case — an event usually checks
    // in against whichever cycle is currently running), then fall back to
    // scanning every other sheet so a QR from an older cycle still resolves.
    candidateNames = [activeName].concat(
      ss.getSheets().map(s => s.getName()).filter(n => n !== activeName && n !== ACTIVITY_LOG_SHEET_NAME)
    );
  }

  for (const name of candidateNames) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];
    const msCol = headers.indexOf("Membership No");
    if (msCol === -1) continue;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][msCol]).trim() === code) {
        return { sheet, headers, rowIndex: i + 1, row: data[i] };
      }
    }
  }
  return null;
}

// action=verifyMember (GET) — read-only lookup, does NOT mark attendance.
// Useful for a security desk that wants to confirm someone's identity
// without accidentally checking them in (e.g. double-checking before
// letting someone back in after stepping out).
function handleVerifyMember_(e) {
  const password = e.parameter.password || "";
  const account = requirePermission_(password, "viewData");
  if (!account) return jsonOutput_({ status: "error", message: "Unauthorized" });

  const code = extractMembershipNo_(e.parameter.code || "");
  if (!code) return jsonOutput_({ status: "error", message: "الكود مش مفهوم" });

  const found = findMemberRowInAllCycles_(code, String(e.parameter.form || "").trim());
  if (!found) return jsonOutput_({ status: "not_found" });

  const chkCol = found.headers.indexOf("Checked In At");
  const checkedInAt = chkCol > -1 && found.row[chkCol] ? found.row[chkCol] : null;
  return jsonOutput_({
    status: "success",
    member: rowToMember_(found.headers, found.row),
    checkedInAt: checkedInAt ? (checkedInAt instanceof Date ? checkedInAt.toISOString() : String(checkedInAt)) : null,
  });
}

// action=checkin (POST) — payload: {password, code, formId}. Marks the
// member as arrived RIGHT NOW, unless they're already checked in (returns
// their original check-in time instead of overwriting it — scanning
// someone twice by accident should never look like an error, just a
// heads-up). formId "" (or omitted) checks every form's data, same as
// before multi-form support existed.
function handleCheckin_(payload) {
  const account = requirePermission_(payload.password || "", "viewData");
  if (!account) return jsonOutput_({ status: "error", message: "Unauthorized" });

  const code = extractMembershipNo_(payload.code || "");
  if (!code) return jsonOutput_({ status: "error", message: "الكود مش مفهوم" });

  const found = findMemberRowInAllCycles_(code, String(payload.formId || "").trim());
  if (!found) return jsonOutput_({ status: "not_found" });

  const { sheet, headers, rowIndex, row } = found;
  let chkCol = headers.indexOf("Checked In At");
  if (chkCol === -1) {
    // Sheet hasn't been healed yet (very old sheet nobody's opened since
    // this feature shipped) — heal it now so the write below has somewhere
    // to land, then re-read the header position.
    healSheetHeaders_(sheet);
    chkCol = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf("Checked In At");
    if (chkCol === -1) return jsonOutput_({ status: "error", message: "تعذر تجهيز عمود الحضور — جرب تاني كمان شوية" });
  }

  const member = rowToMember_(headers, row);
  const existing = row[chkCol];
  if (existing) {
    return jsonOutput_({
      status: "already",
      member,
      checkedInAt: existing instanceof Date ? existing.toISOString() : String(existing),
    });
  }

  const now = new Date();
  sheet.getRange(rowIndex, chkCol + 1).setValue(now);
  logActivity_(account.name, "checkin", `سجّل حضور: ${member.name} (${code})`);
  return jsonOutput_({ status: "success", member, checkedInAt: now.toISOString() });
}

// Sequential membership numbers like OSH-000123, persisted in Script
// Properties so it survives across executions AND across cycles (numbering
// keeps counting up even after a new sheet/cycle is started). First call
// bootstraps the counter from however many rows already exist in the active
// sheet, so numbering picks up naturally even if you're migrating from an
// older sheet.
//
// `prefix` defaults to MEMBERSHIP_PREFIX ("OSH") — a form can use its own
// prefix instead (see cfg.membershipPrefix / the "بادئة رقم العضوية" field
// in Settings) so e.g. "HR-000045" and "OSH-000045" can coexist without
// colliding; each prefix keeps its own independent counter.
// Looks across EVERY sheet in the whole spreadsheet (every form, every
// cycle) for a row already carrying this National ID, and returns that
// row's Membership No if found — so the SAME person registering again
// through a DIFFERENT form (or a new cycle of the same form) keeps their
// existing membership number instead of getting a brand-new one each time.
// (isDuplicateNid_ is a separate, narrower check — it only blocks
// re-registering a SECOND time within the same form's CURRENT active
// sheet; this function is what makes an ID persistent ACROSS forms.)
function findExistingMembershipNo_(nationalId) {
  const ss = getSpreadsheet_();
  const sheets = ss.getSheets().filter(sh => sh.getName() !== ACTIVITY_LOG_SHEET_NAME);

  for (const sh of sheets) {
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;

    const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const nidCol = headerRow.indexOf("National ID");
    const msCol = headerRow.indexOf("Membership No");
    if (nidCol === -1 || msCol === -1) continue;

    const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (const row of values) {
      if (String(row[nidCol]).trim() === nationalId) {
        const existing = String(row[msCol] || "").trim();
        if (existing) return existing;
      }
    }
  }
  return null;
}

function generateMembershipNumber_(prefix) {
  const usePrefix = (prefix || MEMBERSHIP_PREFIX).toUpperCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const used = new Set();
    const ss = getSpreadsheet_();
    const numRe = new RegExp("^" + usePrefix + "-(\\d+)$");

    ss.getSheets().forEach(sh => {
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      if (lastRow < 2 || lastCol < 1) return;

      const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      const colIdx = headerRow.indexOf("Membership No");
      if (colIdx === -1) return;

      const values = sh.getRange(2, colIdx + 1, lastRow - 1, 1).getValues();
      values.forEach(row => {
        const m = String(row[0] || "").trim().match(numRe);
        if (m) used.add(Number(m[1]));
      });
    });

    let counter = 1;
    while (used.has(counter)) counter += 1;
    return `${usePrefix}-${String(counter).padStart(6, "0")}`;
  } finally {
    lock.releaseLock();
  }
}


// ---------------------------------------------------------------------------
// Email
// ---------------------------------------------------------------------------

// Uses the form's own custom subject/body if it set one (Settings ▸ نص
// إيميل التأكيد), with {{name}} and {{membershipNo}} placeholders — falls
// back to the built-in generic Arabic text otherwise.
//
// Also embeds a QR code of the membership number directly in the email
// (as an inline cid: image, not a remote link — so it still shows up even
// if the recipient's mail client blocks external images, and works
// offline once the email is downloaded). On event day, the check-in
// scanner (dys_checkin.html) reads this exact code — see
// extractMembershipNo_() in this file, which is prefix-agnostic and pulls
// the membership number out of whatever text the QR encodes.
function sendConfirmationEmail_(p, membershipNo, cfg) {
  try {
    const fill = (s) => s.replace(/\{\{name\}\}/g, p.name || "").replace(/\{\{membershipNo\}\}/g, membershipNo || "");
    const subject = (cfg && cfg.confirmEmailSubject)
      ? fill(cfg.confirmEmailSubject)
      : "تأكيد التسجيل — أسرة صناع الحياة";
    const bodyText = (cfg && cfg.confirmEmailBody)
      ? fill(cfg.confirmEmailBody)
      : `أهلًا ${p.name}،\n\n` +
        `شكرًا لتسجيلك في أسرة صناع الحياة.\n` +
        `رقم عضويتك هو: ${membershipNo}\n\n` +
        `هيتم التواصل معاك قريبًا من فريق اللجنة.\n\n` +
        `تحياتنا،\nفريق أسرة صناع الحياة`;

    // Best-effort: a QR image failure (network hiccup, qrserver.com down)
    // must NEVER block the confirmation email itself from going out.
    const qrBlob = fetchQrCodeBlob_(membershipNo);

    if (!qrBlob) {
      // No QR available — send exactly like before, plain text only.
      sendEmail_(p.email.trim(), subject, bodyText);
      return true;
    }

    qrBlob.setName("checkin-qr.png");
    const htmlBody =
      `<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;font-size:15px;color:#222;line-height:1.8;">` +
      `<p>${bodyText.replace(/\n/g, "<br>")}</p>` +
      `<div style="text-align:center;margin:22px 0;padding:18px;border:2px dashed #1668a0;border-radius:14px;background:#f7fafc;">` +
      `<p style="margin:0 0 10px;font-weight:bold;color:#0e3c5c;">كود الدخول — اعرضه يوم الإيفنت عشان نسجّل حضورك</p>` +
      `<img src="cid:checkinQr" width="220" height="220" alt="QR كود الدخول" style="display:block;margin:0 auto;">` +
      `<p style="margin:12px 0 0;font-size:13px;color:#666;">أو رقم العضوية يدويًا: <strong>${membershipNo}</strong></p>` +
      `</div>` +
      `</div>`;

    sendEmail_(p.email.trim(), subject, bodyText, { htmlBody, inlineImages: { checkinQr: qrBlob } });
    return true;
  } catch (err) {
    console.error("Email send failed:", err);
    return false;
  }
}


// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Admin accounts (password + granular permissions)
// ---------------------------------------------------------------------------
// The dashboard login box only ever asks for ONE password — there's no
// separate username field. Whoever's password matches determines which
// account (and therefore which permissions) they're logged in as.
//
// Every account can always view the registrations table/charts and switch
// between cycles once logged in — that's just what "having a valid
// dashboard password" means. On top of that, each account is independently
// granted zero or more of these capabilities (see PERMISSION_KEYS at the
// top of the file):
//   manageSettings     — form title, sheet name, registration window, logo,
//                         cycles, diagnostics
//   manageFields       — which form fields are shown/required
//   manageCertificates — upload/remove the cert template, send certificates
//                         (single/bulk/test), toggle auto-send
//   manageAccounts     — add/remove OTHER admin accounts (this is the
//                         sensitive one — see handleRemoveAdminAccount_)
// An account with none of these is a pure read-only "viewer". An account
// with all four behaves like the old "owner" role.
//
// Stored in Script Properties as ADMIN_ACCOUNTS (a JSON array), managed from
// the dashboard's "👥 حسابات الدخول" settings card (needs manageAccounts) —
// you don't need to edit this file or Script Properties by hand to add
// someone.
//
// Backward compatibility: your original single ADMIN_PASSWORD (set via
// setAdminPassword() below) keeps working forever as an implicit
// full-permissions account — you never have to migrate it, it's just always
// included. Any account saved under the OLD role:"owner"/"viewer" shape
// (from before this granular-permissions version) is transparently upgraded
// to the new permissions shape the moment it's read — see normalizeAccount_.

function normalizeAccount_(a) {
  if (a.permissions) {
    // Already the new shape — just re-derive the display-only `role` label.
    return Object.assign({}, a, { role: roleLabelFromPermissions_(a.permissions) });
  }
  // Old shape: {name, password, role: "owner"|"viewer"}.
  const isOwner = a.role === "owner";
  const permissions = {};
  PERMISSION_KEYS.forEach(key => { permissions[key] = isOwner; });
  return Object.assign({}, a, { permissions, role: roleLabelFromPermissions_(permissions) });
}

// Purely cosmetic label for the dashboard's accounts table — "owner" if
// every capability is granted, "viewer" if none are, "custom" otherwise.
function roleLabelFromPermissions_(permissions) {
  const values = PERMISSION_KEYS.map(key => !!permissions[key]);
  if (values.every(v => v)) return "owner";
  if (values.every(v => !v)) return "viewer";
  return "custom";
}

function getRawAdminAccounts_() {
  const raw = PropertiesService.getScriptProperties().getProperty("ADMIN_ACCOUNTS");
  try {
    const arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function getAdminAccounts_() {
  const props = PropertiesService.getScriptProperties();
  let accounts = getRawAdminAccounts_();
  const legacy = props.getProperty("ADMIN_PASSWORD");
  if (legacy && !accounts.some(a => a.password === legacy)) {
    accounts = [{ name: "الحساب الأساسي", password: legacy, role: "owner" }].concat(accounts);
  }
  return accounts.map(normalizeAccount_);
}

function saveAdminAccounts_(accounts) {
  PropertiesService.getScriptProperties().setProperty("ADMIN_ACCOUNTS", JSON.stringify(accounts));
}

// Returns the matching account ({name, password, permissions, role}) or null.
// Tries a Google-signed-in SESSION TOKEN first (see issueSession_ below),
// then falls back to a plain legacy password match — so both auth methods
// work through the exact same code path everywhere else in this file (every
// handler just calls requirePermission_(payload.password, ...) and neither
// knows nor cares which kind of credential it actually got).
function authenticate_(passwordOrToken) {
  if (!passwordOrToken) return null;

  const session = getSession_(passwordOrToken);
  if (session) {
    const account = getAdminAccounts_().find(a => a.email && a.email === session.email && a.status !== "pending");
    return account || null; // account may have been removed/un-approved since the session was issued
  }

  return getAdminAccounts_().find(a => a.password && a.password === passwordOrToken) || null;
}

// Central gate for every password-protected endpoint. `permission` is one of
// PERMISSION_KEYS, or the special value "viewData" which just means "any
// valid logged-in account" (every account can view). Returns the matched
// account on success, or null (and already slept 400ms to blunt
// brute-forcing) on failure — callers just do:
//   const account = requirePermission_(pwd, "manageSettings");
//   if (!account) return jsonOutput_({status:"error", message:"Unauthorized"});
function requirePermission_(password, permission) {
  const account = authenticate_(password);
  const ok = account && (permission === "viewData" || account.permissions[permission] === true);
  if (!ok) {
    Utilities.sleep(400);
    return null;
  }
  return account;
}

// ---------------------------------------------------------------------------
// Google Sign-In — session tokens + the sign-in/approval flow
// ---------------------------------------------------------------------------
// Sessions are stored server-side (Script Properties, pruned of expired
// entries on every new sign-in) rather than as a JWT the client could
// tamper with — the token itself is just an opaque random string that means
// nothing outside this lookup.
const SESSIONS_KEY = "GOOGLE_SESSIONS";

function getSession_(token) {
  const raw = PropertiesService.getScriptProperties().getProperty(SESSIONS_KEY);
  let sessions = {};
  try { sessions = raw ? JSON.parse(raw) : {}; } catch (e) { sessions = {}; }
  const s = sessions[token];
  if (!s || s.expiresAt < Date.now()) return null;
  return s;
}

function issueSession_(email) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(SESSIONS_KEY);
  let sessions = {};
  try { sessions = raw ? JSON.parse(raw) : {}; } catch (e) { sessions = {}; }

  const now = Date.now();
  Object.keys(sessions).forEach(t => { if (sessions[t].expiresAt < now) delete sessions[t]; });

  const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  sessions[token] = { email, expiresAt: now + SESSION_TTL_MS };
  props.setProperty(SESSIONS_KEY, JSON.stringify(sessions));
  return token;
}

// Verifies a Google ID token the RIGHT way for Apps Script (no JWT library
// needed): Google's own tokeninfo endpoint checks the cryptographic
// signature server-side and hands back the decoded claims. We still check
// `aud` ourselves (that the token was actually issued for THIS app's Client
// ID, not some other Google Sign-In button elsewhere) and `email_verified`.
function verifyGoogleIdToken_(idToken) {
  if (!idToken) return null;
  try {
    const res = UrlFetchApp.fetch(
      "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return null;
    const data = JSON.parse(res.getContentText());
    if (data.aud !== GOOGLE_CLIENT_ID) return null;
    if (data.email_verified !== "true" && data.email_verified !== true) return null;
    if (Number(data.exp) * 1000 < Date.now()) return null;
    const email = String(data.email || "").toLowerCase().trim();
    if (!email) return null;
    return { email, name: data.name || email };
  } catch (e) {
    return null;
  }
}

// action=googleLogin (POST) — payload: {idToken}. Three possible outcomes:
//   1. Known + approved email → issue a session token, log them straight in.
//   2. Known but still pending → tell them to wait (no repeat email spam).
//   3. Never seen before → create a pending account + email everyone who can
//      approve new access (manageAccounts), so THEY approve/reject with one
//      tap from their inbox (see handleApproveAccess_/handleRejectAccess_).
function handleGoogleLogin_(payload) {
  const info = verifyGoogleIdToken_(payload.idToken);
  if (!info) return jsonOutput_({ status: "error", message: "فشل التحقق من حساب جوجل — جرب تاني." });

  const accounts = getRawAdminAccounts_();
  const existing = accounts.find(a => a.email === info.email);

  if (existing && existing.status !== "pending") {
    const account = normalizeAccount_(existing);
    const token = issueSession_(info.email);
    return jsonOutput_({
      status: "success",
      sessionToken: token,
      account: { name: account.name, email: account.email, permissions: account.permissions },
    });
  }

  if (!existing) {
    const approvalToken = Utilities.getUuid().replace(/-/g, "");
    accounts.push({
      name: info.name,
      email: info.email,
      status: "pending",
      permissions: {},
      requestedAt: new Date().toISOString(),
      approvalToken,
    });
    saveAdminAccounts_(accounts);
    notifyApprovers_(info, approvalToken);
  }
  // else: already pending from an earlier attempt — don't re-send the email
  // every time they retry signing in.

  return jsonOutput_({ status: "pending", message: "طلبك اتبعت لصاحب صلاحية إدارة الحسابات — هتقدر تدخل أول ما يوافق عليك." });
}

// Emails everyone who currently holds "manageAccounts" (only THEY can grant
// new access) with one-tap approve/reject links. Password-only legacy
// accounts have no email on file and are silently skipped here — they can
// still approve from the dashboard's "👥 حسابات الدخول" card itself.
function notifyApprovers_(info, approvalToken) {
  const approvers = getAdminAccounts_().filter(a => a.email && a.status !== "pending" && a.permissions.manageAccounts);
  if (!approvers.length) return; // nobody CAN approve by email — dashboard-side approval still works

  const base = ScriptApp.getService().getUrl();
  const approveUrl = `${base}?action=approveAccess&email=${encodeURIComponent(info.email)}&token=${encodeURIComponent(approvalToken)}`;
  const rejectUrl = `${base}?action=rejectAccess&email=${encodeURIComponent(info.email)}&token=${encodeURIComponent(approvalToken)}`;

  const subject = `طلب دخول جديد للوحة تحكم أسرة صناع الحياة — ${info.name}`;
  const body =
    `فيه حد طلب يدخل لوحة التحكم بحساب جوجل بتاعه:\n\n` +
    `الاسم: ${info.name}\n` +
    `الإيميل: ${info.email}\n\n` +
    `لو عايز توافق: ${approveUrl}\n\n` +
    `لو مش عايز: ${rejectUrl}\n\n` +
    `(الموافقة بتديله دخول بس من غير أي صلاحيات — تقدر تظبطلها بالظبط بعد كده من "👥 حسابات الدخول" في الإعدادات.)`;

  approvers.forEach(a => {
    try { sendEmail_(a.email, subject, body); } catch (e) { /* one bad address shouldn't block the others */ }
  });
}

// action=approveAccess / action=rejectAccess (GET, clicked from the email —
// NOT a JSON endpoint, since a human opens this straight from their inbox).
// Secured by the per-request approvalToken (not by password), matched
// against the exact pending record — a stale or reused link just fails
// quietly with a plain confirmation page either way.
function handleReviewAccess_(e, approve) {
  const email = String(e.parameter.email || "").toLowerCase().trim();
  const token = String(e.parameter.token || "").trim();
  const accounts = getRawAdminAccounts_();
  const idx = accounts.findIndex(a => a.email === email && a.status === "pending" && a.approvalToken === token);

  let message;
  if (idx === -1) {
    message = "اللينك ده مش شغال — يمكن الطلب اتراجع خلاص، أو اللينك قديم.";
  } else if (approve) {
    delete accounts[idx].status;
    delete accounts[idx].approvalToken;
    saveAdminAccounts_(accounts);
    message = `تمت الموافقة على دخول ${accounts[idx].name || email} ✓ — دلوقتي روح "👥 حسابات الدخول" في الإعدادات وحدد الصلاحيات المناسبة له.`;
  } else {
    accounts.splice(idx, 1);
    saveAdminAccounts_(accounts);
    message = `اترفض طلب الدخول بتاع ${email}.`;
  }

  return HtmlService.createHtmlOutput(
    `<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;max-width:480px;margin:60px auto;padding:24px;border:1px solid #ddd;border-radius:12px;text-align:center;">` +
    `<h2 style="color:#0d2438;">أسرة صناع الحياة</h2><p style="font-size:15px;color:#333;">${message}</p></div>`
  );
}

// Run this ONCE from the Apps Script editor (select it in the function
// dropdown ▸ Run) to set/change the MAIN dashboard password (always an
// "owner" account). Edit the value below first, then run it. To add MORE
// accounts (e.g. a read-only "viewer" for a teammate), use the "👥 حسابات
// الدخول" card in the dashboard instead — you don't need the editor for that.
function setAdminPassword() {
  const NEW_PASSWORD = "OSH.HAYAH"; // <-- edit this line, then Run
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", NEW_PASSWORD);
  Logger.log("Admin password updated.");
}


// ---------------------------------------------------------------------------
// One-time setup / diagnostics — run these manually from the editor as needed
// ---------------------------------------------------------------------------

// Creates the sheet + header row if it doesn't exist yet. Safe to run
// multiple times — it does nothing if the sheet is already there.
function setupSheet() {
  const sheet = getSheet_("");
  Logger.log(`Sheet "${sheet.getName()}" is ready.`);
}

// Compares your sheet's actual row-1 headers against what this script
// expects, and logs any mismatches. Run this BEFORE going live if you're
// pointing this script at a sheet that already has real data in it.
function checkHeaders() {
  const sheet = getSheet_("");
  const actual = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length)).getValues()[0];

  Logger.log("Sheet checked:    " + sheet.getName());
  Logger.log("Expected headers: " + JSON.stringify(HEADERS));
  Logger.log("Actual headers:   " + JSON.stringify(actual));

  const mismatches = [];
  HEADERS.forEach((h, i) => {
    if (String(actual[i] || "").trim() !== h) {
      mismatches.push(`Column ${i + 1}: expected "${h}", found "${actual[i] || "(empty)"}"`);
    }
  });

  if (mismatches.length === 0) {
    Logger.log("✅ Headers match perfectly.");
  } else {
    Logger.log("⚠️ Mismatches found:\n" + mismatches.join("\n"));
    Logger.log("Fix row 1 in the sheet to match 'Expected headers' above before going live.");
  }
}

// Optional convenience: configure the registration window / form title /
// sheet base name directly from the Apps Script editor instead of the
// dashboard's Settings tab. Edit the values below, then Run this once.
// Leave START_AT / END_AT as "" for "no limit" (registration always open).
function configureRegistration_() {
  const FORM_TITLE = "";              // e.g. "تسجيل دفعة 2026" — leave "" to keep the default
  const SHEET_BASE_NAME = "";         // e.g. "Registrations" — leave "" to keep current
  const START_AT = "";                // e.g. "2026-09-01T09:00:00" — leave "" for no limit
  const END_AT = "";                  // e.g. "2026-09-15T23:59:59" — leave "" for no limit
  const START_NEW_CYCLE = false;      // set true to force-create a brand-new sheet right now

  const result = handleSaveConfig_({
    formTitle: FORM_TITLE, sheetBaseName: SHEET_BASE_NAME,
    startAt: START_AT, endAt: END_AT, startNewCycle: START_NEW_CYCLE,
  });
  Logger.log(result.getContent());
}

// Quick self-test you can run from the editor to sanity-check the whole
// pipeline without going through the actual form. Check the Logger output
// (View ▸ Logs) after running.
function testSubmitLocally() {
  const fakePayload = {
    name: "Test User", age: "22", gender: "Male",
    nationalId: "29901011234567", // NOTE: replace with a syntactically valid test ID before running
    phone: "01012345678", whatsapp: "01012345678", email: "",
    address: "", faculty: "Test Faculty", graduate: "No",
    committee: "Volunteer", hasJob: "No", currentJob: "",
    website: "", loadedAt: Date.now() - 5000,
  };
  const result = handleSubmit_(fakePayload);
  Logger.log(result.getContent());
}


// ---------------------------------------------------------------------------
// Output helper
// ---------------------------------------------------------------------------

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


function testEmailNow() {
  try {
    GmailApp.sendEmail("ymahrous05@gmail.com", "اختبار", "ده اختبار إرسال");
    Logger.log("تم الإرسال بنجاح");
  } catch (err) {
    Logger.log("فشل الإرسال: " + err);
  }
}

function checkAiQuotaNow() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  Logger.log("العدد المستخدم النهاردة: " + props.getProperty("AI_QUOTA_" + today));
}

function testGeminiNow() {
  const res = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent",
    {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      headers: { "x-goog-api-key": GEMINI_API_KEY },
      payload: JSON.stringify({
        contents: [{ role: "user", parts: [{ text: "قول أهلاً" }] }],
      }),
    }
  );
  Logger.log("Status: " + res.getResponseCode());
  Logger.log("Body: " + res.getContentText());
}



function testGroqNow() {
  const res = UrlFetchApp.fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + GROK_API_KEY },
    payload: JSON.stringify({
      model: "llama-3.3-70b-versatile",
      messages: [{ role: "user", content: "قول أهلاً" }],
      temperature: 0.4,
      max_tokens: 100,
    }),
  });
  Logger.log("Status: " + res.getResponseCode());
  Logger.log("Body: " + res.getContentText());
}

function testChatNow() {
  const reply = callGrok_(
    "انت مساعد بسيط، رد بجملة واحدة بس.",
    "قول أهلاً",
    []
  );
  Logger.log("Grok reply: " + reply);
}
