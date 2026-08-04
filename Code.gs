// ============================================================
// Code.gs — ০২নং বহেড়াতৈল ইউনিয়ন পরিষদ
// প্রত্যয়নপত্র অটোমেশন সিস্টেম — চূড়ান্ত সংস্করণ (Final)
// ১০০% Google Workspace + Gemini API। কোনো থার্ড-পার্টি সার্ভিস নেই।
//
// এই সংস্করণে যা নতুন/সংশোধিত হয়েছে:
//   ১. seedConfigSheet() — এক ক্লিকে ৪০+ সনদের ধরন Config শীটে বসিয়ে দেয়
//      (আগে Config শীট খালি থাকলে getCertificateTypes() কিছুই রিটার্ন করত না)।
//   ২. generateCertificate()-এ TEMPLATE_DOC_ID/TARGET_FOLDER_ID সেট করা আছে
//      কিনা তা যাচাই + QR তৈরির চারপাশে try/catch, যাতে নেটওয়ার্ক ব্যর্থতায়
//      পুরো সনদ তৈরি না আটকায়।
//   ৩. searchCitizenData: মাস্টার ডিবি না পেলে Certificates শীটে ফলব্যাক করে।
//   ৪. doGet-এ ছোট ছোট স্থিতিশীলতা যাচাই যোগ করা হয়েছে।
// ============================================================

const CONFIG = {
  TEMPLATE_DOC_ID: "PUT_YOUR_TEMPLATE_DOC_ID_HERE",   // আপনার Google Doc টেমপ্লেট ID
  TARGET_FOLDER_ID: "PUT_YOUR_DRIVE_FOLDER_ID_HERE",  // PDF/Doc সংরক্ষণের ফোল্ডার ID
  SHEET_CONFIG: "Config",
  SHEET_CERTIFICATES: "Certificates",
  SHEET_MASTER_DB: "Citizen_Master",
  UNION_NAME: "০২নং বহেড়াতৈল ইউনিয়ন পরিষদ",
  LOCATION: "সখিপুর, টাঙ্গাইল"
};

// ------------------------------------------------------------
// ১. অ্যাডমিন মেনু
// ------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 ইউনিয়ন অটোমেশন')
    .addItem('১️⃣ Certificates ও Master DB শীট প্রস্তুত করুন', 'setupCertificatesSheet')
    .addItem('২️⃣ Config শীট ৪০+ সনদ দিয়ে পূরণ করুন', 'seedConfigSheet')
    .addToUi();
}

function setupCertificatesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(CONFIG.SHEET_CERTIFICATES)) {
    const sheet = ss.insertSheet(CONFIG.SHEET_CERTIFICATES);
    sheet.appendRow([
      "সনদ নম্বর", "ইস্যুর তারিখ", "NID/জন্ম সনদ নম্বর", "আবেদনকারীর নাম",
      "পিতা/স্বামীর নাম", "মাতার নাম", "গ্রাম", "ডাকঘর", "ওয়ার্ড নম্বর",
      "প্রত্যয়নপত্রের ধরন (Key)", "Doc Link", "PDF Link", "Status",
      "Extra Data (JSON)", "মোবাইল"
    ]);
    sheet.setFrozenRows(1);
  }
  if (!ss.getSheetByName(CONFIG.SHEET_MASTER_DB)) {
    const m = ss.insertSheet(CONFIG.SHEET_MASTER_DB);
    m.appendRow(["NID", "নাম", "পিতা/স্বামী", "মাতা", "গ্রাম", "ডাকঘর", "ওয়ার্ড", "লিঙ্গ", "স্বামী/স্ত্রীর নাম", "মোবাইল"]);
    m.setFrozenRows(1);
  }
  SpreadsheetApp.getUi().alert('✅ Certificates ও Master DB শীট প্রস্তুত!');
}

// ------------------------------------------------------------
// ১-ক. Config শীট এক ক্লিকে বীজ-ডেটা দিয়ে পূরণ করা (৪০+ ক্যাটাগরি)
// ------------------------------------------------------------
function seedConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_CONFIG);
  if (sheet && sheet.getLastRow() > 1) {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert('Config শীটে ইতিমধ্যে ডেটা আছে', 'নতুন করে বীজ-ডেটা দিয়ে প্রতিস্থাপন করবেন?', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
    sheet.clear();
  }
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_CONFIG);

  const header = ["Type_Key", "বাংলা_নাম", "Prompt_Hint", "Extra_Simple_Fields", "Extra_Tables"];
  const rows = [
    ["HoldingTax", "হোল্ডিং ট্যাক্স", "বাড়ি/হোল্ডিং করের হালনাগাদ অবস্থা সম্পর্কে লিখো", "holdingNo:হোল্ডিং নং", ""],
    ["FamilyCert", "পারিবারিক সনদ", "পরিবারের সদস্য সংখ্যা ও সম্পর্ক উল্লেখ করে লিখো", "", "members|FAMILY_TABLE|নাম,জন্ম তারিখ,সম্পর্ক"],
    ["Warish", "ওয়ারিশ সনদ", "মৃত ব্যক্তির বৈধ ওয়ারিশগণের কথা উল্লেখ করো", "deceasedName:মৃত ব্যক্তির নাম|deathDate:মৃত্যুর তারিখ", "heirs|HEIRS_TABLE|নাম,জন্ম তারিখ,বয়স,সম্পর্ক,মন্তব্য"],
    ["Inheritance", "উত্তরাধিকার সনদ", "মৃত ব্যক্তির বৈধ উত্তরাধিকারীদের কথা উল্লেখ করো", "deceasedName:মৃত ব্যক্তির নাম|deathDate:মৃত্যুর তারিখ", "heirs|HEIRS_TABLE|নাম,জন্ম তারিখ,বয়স,সম্পর্ক,মন্তব্য"],
    ["PowerOfAttorney", "ক্ষমতা অর্পণ প্রত্যয়ন", "একজন কাউকে ক্ষমতা অর্পণ করার বিষয়ে লিখো", "attorneyName:যাকে ক্ষমতা অর্পণ করা হচ্ছে তার নাম|purpose:ক্ষমতা অর্পণের উদ্দেশ্য", ""],
    ["NewElectricity", "নতুন বিদ্যুৎ সংযোগের প্রত্যয়ন", "নতুন বিদ্যুৎ সংযোগের জন্য উপযুক্ততা সম্পর্কে লিখো", "", ""],
    ["Death", "মৃত্যু প্রত্যয়ন", "ব্যক্তির মৃত্যুর বিষয়ে দাপ্তরিক ভাষায় লিখো", "deathDate:মৃত্যুর তারিখ|deathCause:মৃত্যুর কারণ|deathPlace:মৃত্যুর স্থান", ""],
    ["NotRohingya", "রোহিঙ্গা নয় প্রত্যয়ন", "আবেদনকারী স্থানীয় বাংলাদেশী নাগরিক, রোহিঙ্গা নয় তা উল্লেখ করো", "", ""],
    ["VoterVerification", "ভোটার তথ্য যাচাই-বাছাই প্রত্যয়ন", "ভোটার তথ্য সঠিক ও যাচাইকৃত বলে উল্লেখ করো", "voterNo:ভোটার নং", ""],
    ["Unemployment", "বেকারত্ব সনদ", "আবেদনকারী বর্তমানে বেকার/কর্মহীন বলে উল্লেখ করো", "lastJob:সর্বশেষ পেশা (যদি থাকে)", ""],
    ["Nationality", "জাতীয়তা সনদ", "আবেদনকারী বাংলাদেশী নাগরিক বলে উল্লেখ করো", "", ""],
    ["Citizenship", "নাগরিকত্ব সনদ", "আবেদনকারীর নাগরিকত্বের বিষয়ে লিখো", "", ""],
    ["Remarriage", "পুনঃবিবাহ সনদ", "পুনঃবিবাহের বিষয়ে উল্লেখ করো", "marriageDate:পুনঃবিবাহের তারিখ|spouseName:বর্তমান স্বামী/স্ত্রীর নাম", ""],
    ["Landless", "ভূমিহীন সনদ", "আবেদনকারীর কোনো জমি নেই এই মর্মে লিখো", "", ""],
    ["NewVoter", "নতুন ভোটার প্রত্যয়ন", "নতুন ভোটার তালিকাভুক্তির উপযুক্ততা নিয়ে লিখো", "", ""],
    ["Widow", "বিধবা প্রত্যয়ন", "আবেদনকারী বিধবা এই মর্মে দাপ্তরিক ভাষায় লিখো", "husbandName:স্বামীর নাম|husbandDeathDate:স্বামীর মৃত্যুর তারিখ", ""],
    ["Professional", "পেশাগত সনদ", "আবেদনকারীর পেশা সম্পর্কে উল্লেখ করো", "profession:পেশা|workplace:কর্মস্থল", ""],
    ["Community", "সম্প্রদায় সনদ", "আবেদনকারীর সম্প্রদায়/ধর্মীয় পরিচয় সম্পর্কে লিখো", "community:সম্প্রদায়/ধর্মীয় পরিচয়", ""],
    ["GuardianConsent", "অভিভাবক সম্মতি সনদ", "অভিভাবকের সম্মতি সংক্রান্ত বিষয়ে লিখো", "guardianName:অভিভাবকের নাম|consentPurpose:সম্মতির উদ্দেশ্য", ""],
    ["FreedomFighter", "মুক্তিযোদ্ধা প্রত্যয়ন", "মুক্তিযুদ্ধে অংশগ্রহণ সম্পর্কে সম্মানজনক ভাষায় লিখো", "gazetteNo:মুক্তিযোদ্ধা গেজেট/সনদ নং", ""],
    ["Agriculture", "কৃষি প্রত্যয়ন", "আবেদনকারী কৃষিকাজের সাথে সম্পৃক্ত বলে উল্লেখ করো", "cropType:প্রধান ফসল/কৃষিকাজের ধরন", ""],
    ["Unmarried", "অবিবাহিত সনদ", "আবেদনকারী অবিবাহিত এই মর্মে লিখো", "", ""],
    ["AnnualIncome", "বার্ষিক আয়ের সনদ", "আনুমানিক বার্ষিক আয়ের বিষয়ে সাধারণভাবে লিখো", "annualIncome:বার্ষিক আয় (টাকা)|profession:পেশা", ""],
    ["Orphan", "এতিম সনদ", "আবেদনকারী পিতা-মাতাহীন/এতিম বলে সংবেদনশীল ভাষায় লিখো", "guardianName:বর্তমান অভিভাবকের নাম", ""],
    ["Married", "বিবাহিত প্রত্যয়ন", "আবেদনকারী বিবাহিত এই মর্মে লিখো", "marriageDate:বিবাহের তারিখ|spouseName:স্বামী/স্ত্রীর নাম", ""],
    ["DowryFree", "যৌতুক বিহীন বিবাহ প্রত্যয়ন", "বিবাহে কোনো যৌতুক লেনদেন হয়নি এই মর্মে লিখো", "marriageDate:বিবাহের তারিখ|spouseName:স্বামী/স্ত্রীর নাম", ""],
    ["Passport", "পাসপোর্ট প্রদানের প্রত্যয়ন", "পাসপোর্ট আবেদনের সহায়ক তথ্য হিসেবে লিখো", "purpose:পাসপোর্টের উদ্দেশ্য (ভ্রমণ/চাকরি/অন্যান্য)", ""],
    ["AliveCert", "জীবিত ব্যক্তির প্রত্যয়ন", "আবেদনকারী জীবিত ও সুস্থ আছেন এই মর্মে লিখো", "", ""],
    ["MonthlyIncome", "মাসিক আয়ের সনদ", "আনুমানিক মাসিক আয়ের বিষয়ে সাধারণভাবে লিখো", "monthlyIncome:মাসিক আয় (টাকা)|profession:পেশা", ""],
    ["FinancialSolvency", "আর্থিক স্বচ্ছলতার সনদ", "আবেদনকারী আর্থিকভাবে স্বচ্ছল এই মর্মে লিখো", "profession:পেশা|approxIncome:আনুমানিক আয় (টাকা)", ""],
    ["DeathWhileUnmarried", "অবিবাহিত অবস্থায় মৃত্যু সনদ", "ব্যক্তি অবিবাহিত অবস্থায় মৃত্যুবরণ করেছেন এই মর্মে লিখো", "deathDate:মৃত্যুর তারিখ|deathCause:মৃত্যুর কারণ", ""],
    ["FinancialInsolvency", "আর্থিক অস্বচ্ছলতার সনদ", "আবেদনকারী আর্থিকভাবে অস্বচ্ছল/দরিদ্র এই মর্মে সংবেদনশীল ভাষায় লিখো", "profession:পেশা (যদি থাকে)", ""],
    ["NIDCorrection", "জাতীয় পরিচয় তথ্য সংশোধন প্রত্যয়ন", "পুরাতন ও নতুন তথ্যের পার্থক্য সংশোধনের সুপারিশ আকারে লিখো", "", "corrections|CORRECTION_TABLE|বিষয়,পূর্বের তথ্য,সংশোধিত বাংলা তথ্য,সংশোধিত ইংরেজি তথ্য"],
    ["Childless", "নিঃসন্তান প্রত্যয়ন", "আবেদনকারী নিঃসন্তান এই মর্মে লিখো", "", ""],
    ["Character", "চারিত্রিক সনদ", "আবেদনকারীর চরিত্র ভালো ও কোনো অপরাধমূলক রেকর্ড নেই এই মর্মে লিখো", "", ""],
    ["NoObjection", "অনাপত্তি সনদ", "নির্দিষ্ট বিষয়ে ইউনিয়ন পরিষদের কোনো আপত্তি নেই এই মর্মে লিখো", "subject:কোন বিষয়ে অনাপত্তি", ""],
    ["VoterAreaTransfer", "ভোটার এলাকা স্থানান্তর প্রত্যয়ন", "ভোটার এলাকা স্থানান্তরের সুপারিশ আকারে লিখো", "oldArea:পূর্বের ভোটার এলাকা|newArea:নতুন ভোটার এলাকা", ""],
    ["Disability", "প্রতিবন্ধী সনদ", "আবেদনকারীর শারীরিক/মানসিক প্রতিবন্ধিতা বিষয়ে সম্মানজনক ভাষায় লিখো", "disabilityType:প্রতিবন্ধিতার ধরন|disabilityPercent:প্রতিবন্ধিতার মাত্রা (%)", ""],
    ["SamePerson", "একই ব্যক্তির প্রত্যয়ন", "দুইটি ভিন্ন নথিতে উল্লেখিত ব্যক্তি একই ব্যক্তি এই মর্মে লিখো", "", "names|NAMES_TABLE|পরিচিত নাম,সূত্র/নথি"],
    ["Tribal", "উপজাতি সনদ", "আবেদনকারীর উপজাতি/আদিবাসী পরিচয় সম্পর্কে লিখো", "tribeName:উপজাতি/জাতিগোষ্ঠীর নাম", ""],
    ["PremisesLicense", "প্রিমিসেস লাইসেন্স", "ব্যবসা প্রতিষ্ঠানের প্রিমিসেস লাইসেন্স সংক্রান্ত বিষয়ে লিখো", "businessName:প্রতিষ্ঠানের নাম|businessType:ব্যবসার ধরন", "fees|FEE_TABLE|বিবরণ,পরিমাণ"],
    ["TradeLicense", "ট্রেড লাইসেন্স", "ব্যবসা প্রতিষ্ঠানের ট্রেড লাইসেন্স সংক্রান্ত বিষয়ে লিখো", "businessName:প্রতিষ্ঠানের নাম|businessType:ব্যবসার ধরন", "fees|FEE_TABLE|বিবরণ,পরিমাণ"],
    ["VoterListInclusion", "ভোটার তালিকা বাদ পড়ায় অন্তর্ভুক্তির প্রত্যয়ন", "ভোটার তালিকা থেকে বাদ পড়ার কারণে পুনঃঅন্তর্ভুক্তির সুপারিশ লিখো", "", ""],
    ["NoRemarriage", "পুনঃবিবাহ না হওয়ার প্রত্যয়ন", "আবেদনকারী পুনঃবিবাহ করেননি এই মর্মে লিখো", "", ""],
    ["PermanentResident", "স্থায়ী বাসিন্দা সনদ", "আবেদনকারী এই ইউনিয়নের স্থায়ী বাসিন্দা এই মর্মে লিখো", "residingSince:কত সাল থেকে বসবাস করছেন", ""],
    ["VehicleLicense", "যানবাহন লাইসেন্স", "যানবাহন চলাচলের লাইসেন্স সংক্রান্ত বিষয়ে লিখো", "vehicleType:যানবাহনের ধরন|vehicleNo:যানবাহন নম্বর", "fees|FEE_TABLE|বিবরণ,পরিমাণ"],
    ["Miscellaneous", "বিবিধ সনদ", "সাধারণ প্রশাসনিক প্রত্যয়ন হিসেবে লিখো", "remarks:বিস্তারিত বিবরণ", ""]
  ];

  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, header.length);

  SpreadsheetApp.getUi().alert(`✅ Config শীটে ${rows.length}টি সনদের ধরন যুক্ত করা হয়েছে। প্রয়োজনমতো নতুন সারি যোগ/সম্পাদনা করতে পারবেন।`);
}

// ------------------------------------------------------------
// ২. বাংলা সংখ্যা ও তারিখ হেল্পার
// ------------------------------------------------------------
function toBanglaNum(numStr) {
  if (numStr === null || numStr === undefined) return '';
  const d = { '0': '০', '1': '১', '2': '২', '3': '৩', '4': '৪', '5': '৫', '6': '৬', '7': '৭', '8': '৮', '9': '৯' };
  return numStr.toString().replace(/[0-9]/g, w => d[w]);
}

function toBanglaDate(date) {
  const months = ['জানুয়ারি', 'ফেব্রুয়ারি', 'মার্চ', 'এপ্রিল', 'মে', 'জুন', 'জুলাই', 'আগস্ট', 'সেপ্টেম্বর', 'অক্টোবর', 'নভেম্বর', 'ডিসেম্বর'];
  const d = date || new Date();
  return toBanglaNum(d.getDate()) + ' ' + months[d.getMonth()] + ', ' + toBanglaNum(d.getFullYear());
}

// ------------------------------------------------------------
// ৩. Config শীট থেকে সনদের ধরন লোড করা (ডাটা-চালিত, ৪০+ ক্যাটাগরি)
//    Extra_Simple_Fields: key1:লেবেল১|key2:লেবেল২
//    Extra_Tables:        formKey|TABLE_KEY|হেডার১,হেডার২ ;; (একাধিক টেবিল হলে ;; দিয়ে আলাদা)
// ------------------------------------------------------------
function getCertificateTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_CONFIG);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const types = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const typeKey = row[0].toString().trim();
    types.push({
      key: typeKey,
      label: row[1] ? row[1].toString().trim() : typeKey,
      promptHint: row[2] ? row[2].toString().trim() : '',
      simpleFields: parseSimpleFields(row[3]),
      tables: parseTables(row[4])
    });
  }
  return types;
}

function parseSimpleFields(raw) {
  if (!raw) return [];
  return raw.toString().split('|').filter(Boolean).map(chunk => {
    const parts = chunk.split(':');
    return { key: parts[0].trim(), label: (parts[1] || parts[0]).trim() };
  });
}

function parseTables(raw) {
  if (!raw) return [];
  return raw.toString().split(';;').filter(Boolean).map(chunk => {
    const parts = chunk.split('|');
    const formKey = (parts[0] || '').trim();
    const tableKey = (parts[1] || formKey).trim();
    const headers = (parts[2] || '').split(',').map(h => h.trim()).filter(Boolean);
    return { key: formKey, tableKey: tableKey, headers: headers };
  });
}

function getCertificateTypeByKey(key) {
  return getCertificateTypes().find(t => t.key === key) || null;
}

// ------------------------------------------------------------
// ৪. স্মার্ট সার্চ (মাস্টার ডিবি -> সার্টিফিকেট রেকর্ড ফলব্যাক)
// ------------------------------------------------------------
function searchCitizenData(nid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nidBn = toBanglaNum(nid);

  const master = ss.getSheetByName(CONFIG.SHEET_MASTER_DB);
  if (master) {
    const data = master.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] && (data[i][0].toString() === nid.toString() || data[i][0].toString() === nidBn)) {
        return {
          found: true, name: data[i][1], father: data[i][2], mother: data[i][3],
          village: data[i][4], postOffice: data[i][5], wardNo: data[i][6],
          gender: data[i][7], spouseName: data[i][8], mobile: data[i][9]
        };
      }
    }
  }

  const certs = ss.getSheetByName(CONFIG.SHEET_CERTIFICATES);
  if (certs) {
    const data = certs.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][2] && (data[i][2].toString() === nid.toString() || data[i][2].toString() === nidBn)) {
        return {
          found: true, name: data[i][3], father: data[i][4], mother: data[i][5],
          village: data[i][6], postOffice: data[i][7], wardNo: data[i][8], mobile: data[i][14]
        };
      }
    }
  }
  return { found: false };
}

// ------------------------------------------------------------
// ৫. পুরনো সনদ নম্বর দিয়ে খোঁজা (রিপ্রিন্ট)
// ------------------------------------------------------------
function searchCertificateByNo(certNo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CERTIFICATES);
  if (!sheet) return { found: false };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] && data[i][0].toString().trim() === certNo.toString().trim()) {
      return {
        found: true, name: data[i][3], typeLabel: data[i][9],
        docUrl: data[i][10], pdfUrl: data[i][11],
        docxUrl: driveDocxExportUrl(data[i][10])
      };
    }
  }
  return { found: false };
}

function driveDocxExportUrl(docUrl) {
  if (!docUrl) return '';
  const match = docUrl.toString().match(/[-\w]{25,}/);
  if (!match) return '';
  return `https://docs.google.com/document/d/${match[0]}/export?format=docx`;
}

// ------------------------------------------------------------
// ৬. মাস্টার ডাটাবেসে নাগরিক সংরক্ষণ (নতুন হলেই শুধু যোগ করে)
// ------------------------------------------------------------
function saveToMasterDB(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_MASTER_DB);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_MASTER_DB);
    sheet.appendRow(["NID", "নাম", "পিতা/স্বামী", "মাতা", "গ্রাম", "ডাকঘর", "ওয়ার্ড", "লিঙ্গ", "স্বামী/স্ত্রীর নাম", "মোবাইল"]);
  }
  if (!formData.nid) return; // NID ছাড়া মাস্টার ডিবিতে সংরক্ষণ করা হয় না
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === formData.nid.toString()) return; // ইতিমধ্যে আছে
  }
  sheet.appendRow([
    formData.nid || '', formData.name || '', formData.father || formData.spouseName || '',
    formData.mother || '', formData.village || '', formData.postOffice || '', formData.wardNo || '',
    formData.gender || '', formData.spouseName || '', formData.mobile || ''
  ]);
}

// ------------------------------------------------------------
// ৭. ডুপ্লিকেট-চেক (NID + সনদের ধরন ভিত্তিক)
// ------------------------------------------------------------
function isDuplicateRequest(nid, typeKey) {
  if (!nid) return false;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CERTIFICATES);
  if (!sheet) return false;
  const nidBn = toBanglaNum(nid);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const existingNid = data[i][2] ? data[i][2].toString() : '';
    const existingType = data[i][9] ? data[i][9].toString() : '';
    if ((existingNid === nid.toString() || existingNid === nidBn) && existingType === typeKey) {
      return true;
    }
  }
  return false;
}

// ------------------------------------------------------------
// ৮. সনদ নম্বর জেনারেটর (Auto-increment, বছরভিত্তিক)
// ------------------------------------------------------------
function generateNextCertificateNo(certSheet) {
  const year = new Date().getFullYear();
  const lastRow = certSheet.getLastRow();
  if (lastRow < 2) return `BUP-${year}-00001`;

  const certNos = certSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxSeq = 0;
  const re = new RegExp(`BUP-${year}-(\\d+)`);
  certNos.forEach(v => {
    const s = v ? v.toString() : '';
    const m = s.match(re);
    if (m) maxSeq = Math.max(maxSeq, parseInt(m[1], 10));
  });
  return `BUP-${year}-${String(maxSeq + 1).padStart(5, '0')}`;
}

// ------------------------------------------------------------
// ৯. কোর জেনারেটর — Doc কপি, প্লেসহোল্ডার/টেবিল/QR বসানো, PDF এক্সপোর্ট
// ------------------------------------------------------------
function generateCertificate(formData) {
  if (CONFIG.TEMPLATE_DOC_ID.indexOf('PUT_YOUR') === 0 || CONFIG.TARGET_FOLDER_ID.indexOf('PUT_YOUR') === 0) {
    throw new Error('CONFIG.TEMPLATE_DOC_ID / TARGET_FOLDER_ID এখনো সেট করা হয়নি। Code.gs-এর উপরের দিকে এগুলো বসান।');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let certSheet = ss.getSheetByName(CONFIG.SHEET_CERTIFICATES);
  if (!certSheet) { setupCertificatesSheet(); certSheet = ss.getSheetByName(CONFIG.SHEET_CERTIFICATES); }

  const type = getCertificateTypeByKey(formData.typeKey);
  const typeLabel = type ? type.label : formData.typeKey;
  const promptHint = type ? type.promptHint : '';

  const rawCertNo = generateNextCertificateNo(certSheet);
  const banglaCertNo = toBanglaNum(rawCertNo);
  const issueDate = toBanglaDate(new Date());

  // ক্যাটাগরি-নির্দিষ্ট সাধারণ ফিল্ড — Gemini-কে দেওয়া হবে যাতে বডি-টেক্সটে স্বাভাবিকভাবে বসে
  // (যেমন মৃত্যুর তারিখ/কারণ, হোল্ডিং নং ইত্যাদি), Doc-এ আলাদা করে বসাতে হবে না
  const simpleValues = (formData.extra && formData.extra.simpleFields) || {};

  // Gemini.gs — প্যারামিটার অর্ডার: (name, village, certType, promptHint, extraDetails)
  const bodyText = generateCertificateBodyWithGemini(formData.name, formData.village, typeLabel, promptHint, simpleValues);

  // QR কোড (Google Charts API) — নেটওয়ার্ক ব্যর্থতায় সনদ তৈরি যেন সম্পূর্ণ না আটকায়
  const webAppUrl = ScriptApp.getService().getUrl();
  const verifyUrl = `${webAppUrl}?certNo=${encodeURIComponent(rawCertNo)}`;
  let qrBlob = null;
  try {
    const qrUrl = `https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=${encodeURIComponent(verifyUrl)}&choe=UTF-8`;
    qrBlob = UrlFetchApp.fetch(qrUrl).getBlob();
  } catch (e) {
    Logger.log('QR তৈরি ব্যর্থ: ' + e.toString());
  }

  // Doc টেমপ্লেট কপি
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
  const targetFolder = DriveApp.getFolderById(CONFIG.TARGET_FOLDER_ID);
  const newDoc = templateFile.makeCopy(`${typeLabel} - ${formData.name} (${rawCertNo})`, targetFolder);
  const doc = DocumentApp.openById(newDoc.getId());
  const body = doc.getBody();

  body.replaceText("{{সনদ_নং}}", banglaCertNo);
  body.replaceText("{{ইস্যুর_তারিখ}}", issueDate);
  body.replaceText("{{নাম}}", formData.name || '');
  body.replaceText("{{পিতার_নাম}}", formData.father || formData.spouseName || '');
  body.replaceText("{{মাতার_নাম}}", formData.mother || '');
  body.replaceText("{{গ্রাম}}", formData.village || '');
  body.replaceText("{{ডাকঘর}}", formData.postOffice || '');
  body.replaceText("{{ওয়ার্ড_নং}}", toBanglaNum(formData.wardNo) || '');
  body.replaceText("{{NID_Birth_No}}", toBanglaNum(formData.nid) || '');
  body.replaceText("{{প্রত্যয়নপত্রের_ধরন}}", typeLabel || '');
  body.replaceText("{{body_text}}", bodyText);

  // পৃথক {{key}} প্লেসহোল্ডার থাকলে (ঐচ্ছিক — টেমপ্লেটে সরাসরি না থাকলে কিছু হবে না)
  Object.keys(simpleValues).forEach(k => body.replaceText(`{{${k}}}`, simpleValues[k] || ''));

  const tablesData = (formData.extra && formData.extra.tables) || {};
  if (type && type.tables && type.tables.length) {
    type.tables.forEach(tableDef => {
      const rows = tablesData[tableDef.key] || [];
      insertTableAtPlaceholder(body, `{{TABLE:${tableDef.tableKey}}}`, tableDef.headers, rows);
    });
  }

  if (qrBlob) {
    const qrPlaceholder = body.findText("{{QR_CODE}}");
    if (qrPlaceholder) {
      const element = qrPlaceholder.getElement();
      const parent = element.getParent();
      parent.asParagraph().clear();
      const img = parent.asParagraph().appendImage(qrBlob);
      img.setWidth(110).setHeight(110);
    }
  }

  // QR বসানোর পরে যা-ই {{...}} প্যাটার্নে অবশিষ্ট থাকে (অপ্রাসঙ্গিক টেবিল/ফিল্ড
  // প্লেসহোল্ডার, অথবা QR ব্যর্থ হলে {{QR_CODE}} নিজেই) — সব চুপচাপ মুছে ফেলা হয়
  clearUnusedPlaceholders(body);

  doc.saveAndClose();

  const pdfBlob = newDoc.getAs('application/pdf').setName(`${typeLabel} - ${formData.name}.pdf`);
  const pdfFile = targetFolder.createFile(pdfBlob);

  saveToMasterDB(formData);

  return {
    status: 'success',
    certNo: banglaCertNo,
    issueDate: issueDate,
    url: newDoc.getUrl(),
    docUrl: newDoc.getUrl(),
    pdfUrl: pdfFile.getUrl(),
    docxUrl: driveDocxExportUrl(newDoc.getUrl()),
    bodyText: bodyText,
    typeLabel: typeLabel
  };
}

// একই স্ট্যাটিক টেমপ্লেট ৪০+ ধরনের সনদের জন্য ব্যবহৃত হয়, তাই টেমপ্লেটে এমন {{...}}
// প্লেসহোল্ডারও থাকতে পারে যেগুলো বর্তমান সনদের ধরনে প্রযোজ্য নয় (যেমন Warish-এর
// {{TABLE:HEIRS_TABLE}} একটি Character সনদে অপ্রাসঙ্গিক)। সব নির্দিষ্ট প্রতিস্থাপনের পর
// যা-ই {{...}} প্যাটার্নে অবশিষ্ট থাকে তা চুপচাপ মুছে ফেলা হয়, যাতে প্রিন্ট করা সনদে
// কোনো কাঁচা প্লেসহোল্ডার টেক্সট দেখা না যায়।
function clearUnusedPlaceholders(body) {
  try {
    body.replaceText('\\{\\{[^{}]*\\}\\}', '');
  } catch (e) {
    Logger.log('clearUnusedPlaceholders warning: ' + e.toString());
  }
}

// {{TABLE:key}} প্লেসহোল্ডারকে হেডার-সহ প্রকৃত Doc টেবিল দিয়ে প্রতিস্থাপন করে
function insertTableAtPlaceholder(body, placeholder, headers, rows) {
  const found = body.findText(placeholder.replace(/[{}]/g, '\\$&'));
  if (!found) return;
  const element = found.getElement();
  const parent = element.getParent();
  const index = body.getChildIndex(parent.asParagraph());

  parent.asParagraph().clear();
  if (!rows || !rows.length) return;
  const tableData = [headers].concat(rows);
  body.insertTable(index + 1, tableData);
}

// ------------------------------------------------------------
// ১০. ফর্ম সাবমিট হ্যান্ডলার — সরাসরি সিঙ্ক্রোনাসভাবে সনদ তৈরি করে
// ------------------------------------------------------------
function processCertificateForm(formData) {
  try {
    if (!formData || !formData.typeKey) {
      return { status: 'error', message: 'প্রত্যয়নের ধরন পাওয়া যায়নি।' };
    }
    if (!formData.name || !formData.mother) {
      return { status: 'error', message: 'নাম ও মাতার নাম আবশ্যক।' };
    }
    if (formData.nid && isDuplicateRequest(formData.nid, formData.typeKey)) {
      return {
        status: 'error',
        message: `এই NID (${formData.nid})-এর অধীনে ইতিমধ্যে এই ধরনের একটি সনদ ইস্যু করা হয়েছে। রিপ্রিন্ট করতে সনদ নম্বর দিয়ে খুঁজুন।`
      };
    }

    const result = generateCertificate(formData);

    const certSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CERTIFICATES);
    certSheet.appendRow([
      result.certNo, result.issueDate, toBanglaNum(formData.nid), formData.name,
      formData.father || formData.spouseName || '', formData.mother, formData.village,
      formData.postOffice, toBanglaNum(formData.wardNo), formData.typeKey,
      result.docUrl, result.pdfUrl, "ISSUED",
      JSON.stringify(formData.extra || {}), formData.mobile || ''
    ]);

    return result;
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// ------------------------------------------------------------
// ১১. Gemini Vision গেটওয়ে (NID স্ক্যান)
// ------------------------------------------------------------
function uploadAndParseImages(frontBase64, backBase64) {
  try {
    const raw = extractNidDataWithGemini(frontBase64, backBase64);
    return JSON.parse(raw);
  } catch (e) {
    return { error: 'Gemini প্রতিক্রিয়া পার্স করা যায়নি: ' + e.toString() };
  }
}

// ------------------------------------------------------------
// ১২. Web App এন্ট্রি পয়েন্ট — ফর্ম রেন্ডার + সনদ যাচাইকরণ
// ------------------------------------------------------------
function doGet(e) {
  const certNo = e.parameter.certNo;
  if (!certNo) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle(CONFIG.UNION_NAME + ' — প্রত্যয়নপত্র অটোমেশন')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const certSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CERTIFICATES);
  const data = certSheet ? certSheet.getDataRange().getValues() : [];
  let found = null;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] && data[i][0].toString().trim() === certNo.trim()) {
      found = {
        certNo: data[i][0], issueDate: data[i][1], nid: data[i][2], name: data[i][3],
        father: data[i][4], mother: data[i][5], village: data[i][6], postOffice: data[i][7],
        wardNo: data[i][8], certType: data[i][9]
      };
      break;
    }
  }

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body{font-family:'Segoe UI',sans-serif;background:#eef2f7;margin:0;padding:20px;display:flex;justify-content:center;align-items:center;min-height:100vh;}
    .card{background:#fff;max-width:520px;width:100%;border-radius:12px;box-shadow:0 8px 20px rgba(0,0,0,.1);overflow:hidden;border-top:5px solid #006622;}
    .header{background:#f8fafc;padding:20px;border-bottom:1px solid #eef2f7;text-align:center;}
    .header h2{color:#006622;margin:0;font-size:20px;}
    .badge{background:#28a745;color:#fff;padding:8px 16px;border-radius:20px;font-weight:bold;display:inline-block;margin-top:10px;}
    .content{padding:20px 25px;}
    table{width:100%;border-collapse:collapse;}
    td{padding:8px 6px;border-bottom:1px solid #f0f0f0;font-size:14px;}
    td.l{font-weight:bold;color:#555;width:38%;}
    .err{background:#fff3f3;border-top:5px solid #dc3545;text-align:center;padding:40px;}
  </style></head><body>
  ${found ? `
    <div class="card">
      <div class="header"><h2>🏛️ ${CONFIG.UNION_NAME}</h2><div class="badge">✅ সনদটি সঠিক ও বৈধ</div></div>
      <div class="content"><table>
        <tr><td class="l">সনদ নং</td><td>${toBanglaNum(found.certNo)}</td></tr>
        <tr><td class="l">ধরন</td><td>${found.certType}</td></tr>
        <tr><td class="l">নাম</td><td>${found.name}</td></tr>
        <tr><td class="l">পিতা/স্বামী</td><td>${found.father}</td></tr>
        <tr><td class="l">মাতা</td><td>${found.mother}</td></tr>
        <tr><td class="l">গ্রাম</td><td>${found.village}</td></tr>
        <tr><td class="l">ইস্যুর তারিখ</td><td>${found.issueDate}</td></tr>
      </table></div>
    </div>` : `<div class="card err"><h2 style="color:#dc3545;">❌ সনদটি পাওয়া যায়নি!</h2><p>সনদ নম্বর: ${certNo}</p></div>`}
  </body></html>`;

  return HtmlService.createHtmlOutput(html);
}
