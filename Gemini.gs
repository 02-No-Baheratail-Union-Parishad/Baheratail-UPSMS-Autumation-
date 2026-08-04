// ============================================================
// Gemini.gs — ০২নং বহেড়াতৈল ইউনিয়ন পরিষদ
// Gemini 1.5 Flash ইন্টিগ্রেশন (Vision + Text) — চূড়ান্ত সংস্করণ
// API Key কখনো কোডে হার্ডকোড করা হয় না — Script Properties থেকে পড়া হয়।
// ============================================================

function getGeminiApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error("GEMINI_API_KEY সেট করা নেই। Project Settings > Script Properties এ যোগ করুন।");
  return key;
}

// ১. NID ভিশন পার্সার (সামনে ও পেছনের ছবি থেকে তথ্য বের করা)
function extractNidDataWithGemini(frontBase64, backBase64) {
  let apiKey;
  try { apiKey = getGeminiApiKey(); }
  catch (e) { return JSON.stringify({ error: e.message }); }

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

  const prompt = `
    You are a highly accurate Bengali OCR system for a Bangladesh National ID (NID) card.
    Analyze the provided Front Image and Back Image.
    Extract:
    1. NID Number (10 or 17 digits) - Front or Back.
    2. Applicant Name (in Bangla) - Front.
    3. Father's Name (in Bangla) - Back.
    4. Mother's Name (in Bangla) - Back.
    5. Date of Birth - Back.
    6. Address/Village/Post Office (in Bangla) - Back.

    STRICT RULES:
    - Return ONLY a raw JSON object. No markdown, no code fences, no explanation.
    - If a field is not visible, return "" for that key.
    - JSON keys must be exactly: "nidNo", "name", "fatherName", "motherName", "dob", "addressText".
  `;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: "image/jpeg", data: frontBase64 } },
        { inline_data: { mime_type: "image/jpeg", data: backBase64 } }
      ]
    }],
    generationConfig: { temperature: 0.1 }
  };

  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const code = response.getResponseCode();
    const json = JSON.parse(response.getContentText());

    if (code !== 200) {
      const msg = (json.error && json.error.message) || `HTTP ${code}`;
      return JSON.stringify({ error: "Gemini API ত্রুটি: " + msg });
    }

    const candidate = json.candidates && json.candidates[0];
    if (candidate && candidate.finishReason && candidate.finishReason !== 'STOP') {
      return JSON.stringify({ error: "Gemini নিরাপত্তা/দৈর্ঘ্য কারণে সম্পূর্ণ উত্তর দেয়নি (" + candidate.finishReason + ")। ছবি আবার তুলে চেষ্টা করুন।" });
    }

    if (candidate && candidate.content && candidate.content.parts[0].text) {
      let text = candidate.content.parts[0].text.trim();
      text = text.replace(/```json/gi, '').replace(/```/g, '').trim();
      JSON.parse(text); // বৈধতা যাচাই — অবৈধ হলে catch ব্লকে যাবে
      return text;
    }
    return JSON.stringify({ error: "Gemini কোনো বৈধ প্রতিক্রিয়া দেয়নি।" });
  } catch (e) {
    return JSON.stringify({ error: "Gemini প্রতিক্রিয়া প্রক্রিয়াকরণে ত্রুটি: " + e.toString() });
  }
}

// ২. দাপ্তরিক বাংলা বডি টেক্সট জেনারেটর
// প্যারামিটার অর্ডার: (applicantName, village, certType, promptHint, extraDetails)
// — Code.gs-এর generateCertificate() ঠিক এই অর্ডারেই কল করে।
// extraDetails: Config শীটের Extra_Simple_Fields থেকে আসা key:value অবজেক্ট
// (যেমন { deathDate: "১৫-০৩-২০২৬", deathCause: "স্বাভাবিক" }) — এগুলো এখন সরাসরি
// Doc প্লেসহোল্ডারে না বসিয়ে Gemini-কে দেওয়া হয়, যাতে বডি টেক্সটেই স্বাভাবিকভাবে বসে।
function generateCertificateBodyWithGemini(applicantName, village, certType, promptHint, extraDetails) {
  let apiKey;
  try { apiKey = getGeminiApiKey(); }
  catch (e) { return getDefaultBodyText(applicantName, village, certType); }

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

  const hintLine = promptHint ? `বিশেষ নির্দেশনা: ${promptHint}` : '';

  let extraLines = '';
  if (extraDetails && typeof extraDetails === 'object') {
    const keys = Object.keys(extraDetails).filter(k => extraDetails[k]);
    if (keys.length) {
      extraLines = 'অতিরিক্ত তথ্য (প্রাসঙ্গিক হলে বডি টেক্সটে স্বাভাবিকভাবে বসাও):\n' +
        keys.map(k => `- ${k}: ${extraDetails[k]}`).join('\n');
    }
  }

  const prompt = `
    তুমি ${CONFIG.UNION_NAME}, ${CONFIG.LOCATION}-এর একজন দক্ষ ও পেশাদার প্রশাসনিক লেখক।
    আবেদনকারীর নাম: ${applicantName}
    গ্রাম: ${village}
    প্রত্যয়নপত্রের ধরন: ${certType}
    ${hintLine}
    ${extraLines}

    নির্দেশনা (Zero-Fluff Rule):
    ১. "এই মর্মে প্রত্যয়ন করা যাচ্ছে যে..." দিয়ে শুরু করে ৪-৫ লাইনের একটি সুসংগঠিত দাপ্তরিক প্যারাগ্রাফ লিখো।
    ২. কোনো ভূমিকা, সম্ভাষণ বা শুভেচ্ছাবাণী লেখা যাবে না — শুধু মূল দাপ্তরিক বর্ণনা।
    ৩. ভাষা অত্যন্ত মার্জিত, আনুষ্ঠানিক ও প্রশাসনিক বাংলা হবে।
    ৪. উপরে দেওয়া "অতিরিক্ত তথ্য" থাকলে সেগুলো বাক্যের ভেতরে স্বাভাবিকভাবে বুনে দাও (যেমন তারিখ/কারণ/সম্পর্ক), আলাদা তালিকা বা লেবেল হিসেবে নয়।
    ৫. যা দেওয়া হয়নি এমন কোনো তারিখ/সংখ্যা/তথ্য নিজে থেকে বানিয়ে বলবে না।
  `;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.4 }
  };
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('Gemini text API non-200: ' + response.getContentText());
      return getDefaultBodyText(applicantName, village, certType);
    }
    const json = JSON.parse(response.getContentText());
    const candidate = json.candidates && json.candidates[0];
    if (candidate && candidate.content && candidate.content.parts[0].text) {
      return candidate.content.parts[0].text.trim();
    }
  } catch (e) {
    Logger.log('Gemini text generation error: ' + e.toString());
  }

  return getDefaultBodyText(applicantName, village, certType);
}

// ৩. Gemini ব্যর্থ হলে বা কী সেট না থাকলে ফলব্যাক টেক্সট
function getDefaultBodyText(name, village, certType) {
  return `এই মর্মে প্রত্যয়ন করা যাচ্ছে যে, ${name} ${CONFIG.UNION_NAME}-এর ${village || ''} গ্রামের একজন স্থায়ী বাসিন্দা। তিনি "${certType}" পাওয়ার জন্য আইনানুগভাবে যোগ্য বলে বিবেচিত হয়েছেন। তাহার আচার-আচরণ ও সার্বিক তথ্য পরিষদের নথিতে সন্তোষজনক বলে পরিলক্ষিত হয়েছে।`;
}
