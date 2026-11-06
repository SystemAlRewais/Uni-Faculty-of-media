
/**
 * نظام الإدارة الأكاديمية السحابي - كلية الإعلام
 * Google Apps Script Backend
 * شغّل setupSystem() مرة واحدة قبل النشر.
 */

const SHEETS = {
  students: ["الطلاب", [
    "الكود الجامعي (ID)", "الاسم", "الهاتف", "الفرقة", "كلمة المرور"
  ]],
  content: ["محتوى_الفرقة", [
    "الفرقة", "النوع", "العنوان", "الرابط"
  ]],
  settings: ["إعدادات_المحاضرة", [
    "الفرقة", "اسم المادة", "رابط الخريطة", "النطاق بالمتر",
    "حالة الفتح", "الكود السري للمحاضرة", "خط العرض (Lat)", "خط الطول (Lng)"
  ]],
  attendance: ["الحضور_الحي", [
    "الوقت والتاريخ", "الفرقة", "كود الطالب", "اسم الطالب", "المادة"
  ]],
  archive: ["الأرشيف", [
    "الفرقة", "اسم المادة", "تاريخ الأرشيف", "عدد الحضور", "رابط ملف CSV"
  ]],
  announcements: ["الإعلانات", [
    "الوقت والتاريخ", "الفرقة", "نص الإعلانات"
  ]],
  complaints: ["الشكاوى", [
    "الوقت والتاريخ", "الفرقة", "كود الطالب", "اسم الطالب", "نص الشكوى"
  ]]
};

const SESSION_TTL = 21600;
const INITIAL_ADMIN_USER = "admin";
const INITIAL_ADMIN_PASSWORD = "44554455";


/* ==============================
   تشغيل واجهة الموقع
================================ */

function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("النظام الأكاديمي | كلية الإعلام")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}


/* ==============================
   إنشاء قاعدة البيانات
================================ */

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      "افتح Apps Script من ملف Google Sheets المرتبط أولاً."
    );
  }

  PropertiesService.getScriptProperties()
    .setProperty("DB_ID", ss.getId());

  Object.keys(SHEETS).forEach(key => {
    const [name, headers] = SHEETS[key];

    let sh = ss.getSheetByName(name);

    if (!sh) {
      sh = ss.insertSheet(name);
    }

    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
    }

    sh.setFrozenRows(1);

    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold");
  });

  const props = PropertiesService.getScriptProperties();

  if (!props.getProperty("ADMIN_HASH")) {
    const salt = Utilities.getUuid();

    props.setProperties({
      ADMIN_USER: INITIAL_ADMIN_USER,
      ADMIN_SALT: salt,
      ADMIN_HASH: hashPassword_(
        INITIAL_ADMIN_PASSWORD,
        salt
      )
    });
  }

  return "تم تجهيز النظام والتبويبات وحساب الإدارة الأولي.";
}


/* ==============================
   وظائف مساعدة
================================ */

function safeError_(e) {
  const msg = String(
    e && e.message || "حدث خطأ غير متوقع"
  );

  return msg.length > 180
    ? "تعذر تنفيذ الطلب. راجع الإعدادات وحاول مجدداً."
    : msg;
}

function db_() {
  const id = PropertiesService
    .getScriptProperties()
    .getProperty("DB_ID");

  if (!id) {
    throw new Error("شغّل setupSystem() أولاً.");
  }

  return SpreadsheetApp.openById(id);
}

function sheet_(key) {
  return db_().getSheetByName(SHEETS[key][0]);
}

function rows_(key) {
  const sh = sheet_(key);
  const vals = sh.getDataRange().getValues();

  return vals.length < 2 ? [] : vals.slice(1);
}

function clean_(v) {
  return String(v == null ? "" : v).trim();
}

function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + "|" + clean_(password),
    Utilities.Charset.UTF_8
  );

  return bytes.map(b =>
    ("0" + ((b + 256) % 256).toString(16)).slice(-2)
  ).join("");
}

function newPasswordRecord_(password) {
  const salt = Utilities.getUuid();

  return salt + ":" + hashPassword_(password, salt);
}

function verifyPassword_(password, record) {
  const p = clean_(record).split(":");

  return p.length === 2 &&
    hashPassword_(password, p[0]) === p[1];
}


/* ==============================
   إدارة الجلسات
================================ */

function makeSession_(payload) {
  const token = Utilities.getUuid() + Utilities.getUuid();

  CacheService.getScriptCache().put(
    "sess_" + token,
    JSON.stringify(payload),
    SESSION_TTL
  );

  return token;
}

function session_(token, role) {
  const raw = CacheService.getScriptCache()
    .get("sess_" + clean_(token));

  if (!raw) {
    throw new Error("انتهت الجلسة. سجّل الدخول مرة أخرى.");
  }

  const s = JSON.parse(raw);

  if (role && s.role !== role) {
    throw new Error("غير مصرح بتنفيذ هذه العملية.");
  }

  return s;
}


/* ==============================
   تسجيل الدخول والخروج
================================ */

function loginStudent(code, password) {
  code = clean_(code);
  password = clean_(password);

  const found = rows_("students").find(
    r => clean_(r[0]) === code
  );

  if (!found || !verifyPassword_(password, found[4])) {
    throw new Error("بيانات الدخول غير صحيحة.");
  }

  const user = {
    role: "student",
    id: clean_(found[0]),
    name: clean_(found[1]),
    phone: clean_(found[2]),
    grade: clean_(found[3])
  };

  return {
    token: makeSession_(user),
    user: user
  };
}

function loginAdmin(username, password) {
  const p = PropertiesService.getScriptProperties();
  const salt = p.getProperty("ADMIN_SALT") || "";

  if (
    clean_(username) !== p.getProperty("ADMIN_USER") ||
    hashPassword_(password, salt) !== p.getProperty("ADMIN_HASH")
  ) {
    throw new Error("بيانات الدخول غير صحيحة.");
  }

  const user = {
    role: "admin",
    name: "الإدارة"
  };

  return {
    token: makeSession_(user),
    user: user
  };
}

function logout(token) {
  CacheService.getScriptCache()
    .remove("sess_" + clean_(token));

  return true;
}


/* ==============================
   لوحة الطالب
================================ */

function getStudentDashboard(token) {
  const u = session_(token, "student");

  const content = rows_("content")
    .filter(r => clean_(r[0]) === u.grade)
    .map(r => ({
      type: clean_(r[1]),
      title: clean_(r[2]),
      url: clean_(r[3])
    }));

  const announcements = rows_("announcements")
    .filter(r =>
      clean_(r[1]) === u.grade ||
      clean_(r[1]) === "الكل"
    )
    .slice(-30)
    .reverse()
    .map(r => ({
      date: r[0],
      text: clean_(r[2])
    }));

  const lecture = rows_("settings").find(
    r =>
      clean_(r[0]) === u.grade &&
      String(r[4]).toLowerCase() === "true"
  );

  return {
    user: u,
    content: content,
    announcements: announcements,

    lecture: lecture
      ? {
          subject: clean_(lecture[1]),
          map: clean_(lecture[2]),
          radius: Number(lecture[3]),
          open: true
        }
      : null
  };
}


/* ==============================
   بيانات لوحة الإدارة
================================ */

function adminGetData(token, grade) {
  session_(token, "admin");

  grade = clean_(grade);

  const students = rows_("students")
    .filter(r => !grade || clean_(r[3]) === grade)
    .map(r => ({
      id: clean_(r[0]),
      name: clean_(r[1]),
      phone: clean_(r[2]),
      grade: clean_(r[3])
    }));

  const contentValues = sheet_("content")
    .getDataRange()
    .getValues();

  const content = contentValues
    .slice(1)
    .map((r, i) => ({
      r: r,
      row: i + 2
    }))
    .filter(x => !grade || clean_(x.r[0]) === grade)
    .map(x => ({
      row: x.row,
      grade: clean_(x.r[0]),
      type: clean_(x.r[1]),
      title: clean_(x.r[2]),
      url: clean_(x.r[3])
    }));

  const settings = rows_("settings")
    .filter(r => !grade || clean_(r[0]) === grade)
    .map(r => ({
      grade: clean_(r[0]),
      subject: clean_(r[1]),
      map: clean_(r[2]),
      radius: r[3],
      open: String(r[4]).toLowerCase() === "true",
      code: clean_(r[5]),
      lat: r[6],
      lng: r[7]
    }));

  const attendance = rows_("attendance")
    .filter(r => !grade || clean_(r[1]) === grade)
    .slice(-200)
    .reverse()
    .map(r => ({
      time: r[0],
      grade: clean_(r[1]),
      id: clean_(r[2]),
      name: clean_(r[3]),
      subject: clean_(r[4])
    }));

  const archive = rows_("archive")
    .filter(r => !grade || clean_(r[0]) === grade)
    .slice(-100)
    .reverse()
    .map(r => ({
      grade: clean_(r[0]),
      subject: clean_(r[1]),
      date: r[2],
      count: r[3],
      url: clean_(r[4])
    }));

  const complaints = rows_("complaints")
    .filter(r => !grade || clean_(r[1]) === grade)
    .slice(-100)
    .reverse()
    .map(r => ({
      time: r[0],
      grade: clean_(r[1]),
      id: clean_(r[2]),
      name: clean_(r[3]),
      text: clean_(r[4])
    }));

  return {
    students,
    content,
    settings,
    attendance,
    archive,
    complaints
  };
}


/* ==============================
   إدارة الطلاب
================================ */

function registerStudent(token, data) {
  session_(token, "admin");

  data = data || {};

  const name = clean_(data.name);
  const phone = clean_(data.phone);
  const grade = clean_(data.grade);
  const password = clean_(data.password);

  if (!name || !grade || password.length < 6) {
    throw new Error(
      "أدخل الاسم والفرقة وكلمة مرور من 6 أحرف على الأقل."
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const id = "AZ" +
      Utilities.getUuid()
        .replace(/-/g, "")
        .substring(0, 6)
        .toUpperCase();

    sheet_("students").appendRow([
      id,
      name,
      phone,
      grade,
      newPasswordRecord_(password)
    ]);

    return {
      id,
      name,
      phone,
      grade
    };

  } finally {
    lock.releaseLock();
  }
}

function updateStudent(token, data) {
  session_(token, "admin");

  data = data || {};

  const id = clean_(data.id);
  const sh = sheet_("students");
  const vals = sh.getDataRange().getValues();

  const row = vals.findIndex(
    (r, i) => i > 0 && clean_(r[0]) === id
  );

  if (row < 1) {
    throw new Error("الطالب غير موجود.");
  }

  sh.getRange(row + 1, 2, 1, 3).setValues([[
    clean_(data.name),
    clean_(data.phone),
    clean_(data.grade)
  ]]);

  if (clean_(data.password)) {
    if (clean_(data.password).length < 6) {
      throw new Error(
        "كلمة المرور يجب ألا تقل عن 6 أحرف."
      );
    }

    sh.getRange(row + 1, 5)
      .setValue(newPasswordRecord_(data.password));
  }

  return true;
}

function deleteStudent(token, id) {
  session_(token, "admin");

  const sh = sheet_("students");
  const vals = sh.getDataRange().getValues();

  const row = vals.findIndex(
    (r, i) => i > 0 && clean_(r[0]) === clean_(id)
  );

  if (row < 1) {
    throw new Error("الطالب غير موجود.");
  }

  sh.deleteRow(row + 1);

  return true;
}


/* ==============================
   إدارة المحتوى الدراسي
================================ */

function addGradeContent(token, data) {
  session_(token, "admin");

  data = data || {};

  const grade = clean_(data.grade);
  const type = clean_(data.type);
  const title = clean_(data.title);
  const url = clean_(data.url);

  if (
    !grade ||
    !["book", "schedule", "exam"].includes(type) ||
    !title ||
    !/^https?:\/\//i.test(url)
  ) {
    throw new Error(
      "تحقق من الفرقة والنوع والعنوان والرابط."
    );
  }

  sheet_("content").appendRow([
    grade,
    type,
    title,
    url
  ]);

  return true;
}

function deleteGradeContent(token, row) {
  session_(token, "admin");

  row = Number(row);

  if (row < 2) {
    throw new Error("عنصر غير صالح.");
  }

  sheet_("content").deleteRow(row);

  return true;
}


/* ==============================
   إعداد المحاضرات
================================ */

function saveLectureSettings(token, data) {
  session_(token, "admin");

  data = data || {};

  const grade = clean_(data.grade);
  const subject = clean_(data.subject);
  const map = clean_(data.map);
  const radius = Number(data.radius);
  const code = clean_(data.code);
  const lat = Number(data.lat);
  const lng = Number(data.lng);

  if (
    !grade ||
    !subject ||
    !code ||
    !Number.isFinite(radius) ||
    radius < 10 ||
    radius > 2000 ||
    !Number.isFinite(lat) ||
    !Number.isFinite(lng) ||
    lat < -90 ||
    lat > 90 ||
    lng < -180 ||
    lng > 180
  ) {
    throw new Error(
      "بيانات المحاضرة أو الإحداثيات غير صحيحة."
    );
  }

  const sh = sheet_("settings");
  const vals = sh.getDataRange().getValues();

  const idx = vals.findIndex(
    (r, i) => i > 0 && clean_(r[0]) === grade
  );

  const row = [
    grade,
    subject,
    map,
    radius,
    false,
    code,
    lat,
    lng
  ];

  if (idx > 0) {
    sh.getRange(idx + 1, 1, 1, 8)
      .setValues([row]);
  } else {
    sh.appendRow(row);
  }

  return true;
}

function toggleAttendance(token, grade, open) {
  session_(token, "admin");

  const sh = sheet_("settings");
  const vals = sh.getDataRange().getValues();

  const idx = vals.findIndex(
    (r, i) => i > 0 && clean_(r[0]) === clean_(grade)
  );

  if (idx < 1) {
    throw new Error("احفظ إعدادات المحاضرة أولاً.");
  }

  sh.getRange(idx + 1, 5).setValue(!!open);

  return true;
}


/* ==============================
   حساب المسافة GPS
================================ */

function haversine_(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const rad = x => x * Math.PI / 180;

  const dLat = rad(lat2 - lat1);
  const dLon = rad(lon2 - lon1);

  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(rad(lat1)) *
    Math.cos(rad(lat2)) *
    Math.sin(dLon / 2) ** 2;

  return 2 * R *
    Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}


/* ==============================
   تسجيل الحضور
================================ */

function recordAttendance(token, code, lat, lng) {
  const u = session_(token, "student");

  const settings = rows_("settings").find(
    r =>
      clean_(r[0]) === u.grade &&
      String(r[4]).toLowerCase() === "true"
  );

  if (!settings) {
    throw new Error(
      "لا توجد محاضرة مفتوحة لفرقتك حالياً."
    );
  }

  if (clean_(code) !== clean_(settings[5])) {
    throw new Error("الكود السري غير صحيح.");
  }

  lat = Number(lat);
  lng = Number(lng);

  if (
    !Number.isFinite(lat) ||
    !Number.isFinite(lng)
  ) {
    throw new Error("تعذر قراءة موقعك.");
  }

  const distance = haversine_(
    lat,
    lng,
    Number(settings[6]),
    Number(settings[7])
  );

  if (distance > Number(settings[3])) {
    throw new Error(
      "أنت خارج النطاق المسموح. المسافة التقريبية: " +
      Math.round(distance) + " متر."
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const existing = rows_("attendance").some(
      r =>
        clean_(r[1]) === u.grade &&
        clean_(r[2]) === u.id &&
        clean_(r[4]) === clean_(settings[1])
    );

    if (existing) {
      throw new Error(
        "تم تسجيل حضورك لهذه المحاضرة بالفعل."
      );
    }

    sheet_("attendance").appendRow([
      new Date(),
      u.grade,
      u.id,
      u.name,
      clean_(settings[1])
    ]);

    return {
      success: true,
      subject: clean_(settings[1])
    };

  } finally {
    lock.releaseLock();
  }
}


/* ==============================
   أرشفة الحضور CSV
================================ */

function archiveLecture(token, grade) {
  session_(token, "admin");

  grade = clean_(grade);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const settings = rows_("settings").find(
      r => clean_(r[0]) === grade
    );

    if (!settings) {
      throw new Error(
        "لا توجد إعدادات لهذه الفرقة."
      );
    }

    const subject = clean_(settings[1]);
    const sh = sheet_("attendance");
    const all = sh.getDataRange().getValues();

    const matches = all.slice(1).filter(
      r =>
        clean_(r[1]) === grade &&
        clean_(r[4]) === subject
    );

    if (!matches.length) {
      throw new Error(
        "لا توجد سجلات حضور لأرشفتها."
      );
    }

    const csv = [
      [
        "الوقت والتاريخ",
        "الفرقة",
        "كود الطالب",
        "اسم الطالب",
        "المادة"
      ],
      ...matches
    ]
      .map(row => row.map(csvCell_).join(","))
      .join("\r\n");

    const blob = Utilities.newBlob(
      "\uFEFF" + csv,
      "text/csv",
      grade + "_" +
      subject + "_" +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyyMMdd_HHmm"
      ) + ".csv"
    );

    const file = DriveApp.createFile(blob);

    sheet_("archive").appendRow([
      grade,
      subject,
      new Date(),
      matches.length,
      file.getUrl()
    ]);

    for (let i = all.length - 1; i >= 1; i--) {
      if (
        clean_(all[i][1]) === grade &&
        clean_(all[i][4]) === subject
      ) {
        sh.deleteRow(i + 1);
      }
    }

    const settingsRows = rows_("settings");
    const settingsIndex = settingsRows.findIndex(
      r => clean_(r[0]) === grade
    );

    if (settingsIndex >= 0) {
      sheet_("settings")
        .getRange(settingsIndex + 2, 5)
        .setValue(false);
    }

    return {
      count: matches.length,
      url: file.getUrl()
    };

  } finally {
    lock.releaseLock();
  }
}

function csvCell_(v) {
  return '"' +
    String(v == null ? "" : v).replace(/"/g, '""') +
    '"';
}


/* ==============================
   الإعلانات
================================ */

function addAnnouncement(token, grade, message) {
  session_(token, "admin");

  message = clean_(message);

  if (!message) {
    throw new Error("اكتب نص الإعلان.");
  }

  sheet_("announcements").appendRow([
    new Date(),
    clean_(grade) || "الكل",
    message
  ]);

  return true;
}


/* ==============================
   الشكاوى والمقترحات
================================ */

function submitComplaint(token, message) {
  const u = session_(token, "student");

  message = clean_(message);

  if (message.length < 5) {
    throw new Error(
      "اكتب تفاصيل الشكوى أو المقترح."
    );
  }

  sheet_("complaints").appendRow([
    new Date(),
    u.grade,
    u.id,
    u.name,
    message
  ]);

  return true;
}


/* ==============================
   تغيير كلمة مرور الإدارة
================================ */

function changeAdminPassword(token, newPassword) {
  session_(token, "admin");

  newPassword = clean_(newPassword);

  if (newPassword.length < 10) {
    throw new Error(
      "استخدم كلمة مرور من 10 أحرف على الأقل."
    );
  }

  const salt = Utilities.getUuid();
  const p = PropertiesService.getScriptProperties();

  p.setProperties({
    ADMIN_SALT: salt,
    ADMIN_HASH: hashPassword_(newPassword, salt)
  });

  return true;
}
