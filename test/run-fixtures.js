/* =====================================================================
   اختبار النواة على الحالات المأخوذة من محرك InDesign العامل.
   يُشغَّل بلا أي اعتماد خارجي:   node test/run-fixtures.js
   الاجتياز الكامل شرط القبول.
   ===================================================================== */

var fs = require("fs");
var path = require("path");
var core = require(path.join(__dirname, "..", "src", "core.js"));

var fixtures = JSON.parse(
    fs.readFileSync(path.join(__dirname, "fixtures.json"), "utf8"));

var passed = 0;
var failures = [];

function show(text) {
    return JSON.stringify(text).slice(1, -1);
}

fixtures.forEach(function (f, i) {
    var converted = core.convertText(f.input);
    var restored = core.restoreText(converted);
    var problems = [];

    if (converted !== f.convert) {
        problems.push("التحويل:\n      المنتظر: " + show(f.convert) +
                      "\n      الناتج : " + show(converted));
    }
    if (restored !== f.restore) {
        problems.push("الاسترجاع:\n      المنتظر: " + show(f.restore) +
                      "\n      الناتج : " + show(restored));
    }

    if (problems.length) {
        failures.push("  [" + i + "] " + f.label + "\n      المدخل : " +
                      show(f.input) + "\n    " + problems.join("\n    "));
    } else {
        passed++;
    }
});

/* اختبارات إضافية لا تشملها الحالات: خيار المسافة والرموز المكررة. */
var extra = [];

function check(label, actual, expected) {
    if (actual !== expected) {
        extra.push("  " + label + "\n      المنتظر: " + show(expected) +
                   "\n      الناتج : " + show(actual));
    } else {
        passed++;
    }
}

check("المسافة غير المنقسمة قبل الرمز",
      core.convertText("قال النبي صلى الله عليه وآله وسلم كلاماً",
                       { spaceChar: " " }),
      "قال النبي M كلاماً");

check("استرجاع يعيد المسافة الخاصة عادية",
      core.restoreText("قال النبي M كلاماً"),
      "قال النبي صلى الله عليه وآله كلاماً");

/*
    النواة لا تعرف التنسيق، فكل حرف من حروف الرموز مرشَّح للاسترجاع.
    وتمييزُ الرمز الحقيقي من حرف لاتيني عادي في نصّ إنجليزي وظيفةُ
    طبقة Word: تفحص نمط الحرف واسم الخط قبل الإحلال. والمفحوص هنا
    أن الترتيب يُحسب صحيحاً ليصح تزاوج المطابقات مع نطاقات البحث.
*/
check("كلمة إنجليزية لا تُقرأ رموزاً",
      core.restoreText("Word processing"), "Word processing");
check("رمزان متجاوران لا يسقطان",
      core.restoreText("TG"), "عليه السلامرحمه الله");

var restorePlan = core.planRestore("قال T ثم T مرة أخرى");
check("عدد الرموز المرشحة", String(restorePlan.ops.length), "2");
check("ترتيب الرمز الثاني",
      String(restorePlan.ops[1].symbolOrdinal), "1");
check("ترتيب النصّ الحرفي داخل الفقرة",
      core.literalOccurrences("ب ب ب", "ب").join(","), "0,2,4");

check("تعطيل حذف الشرطتين",
      core.convertText("الشيخ -قدس سره- كتب", { stripDashes: false }),
      "الشيخ -P- كتب");

check("تعطيل فحص السياق يحوّل المحجوب",
      core.convertText("رد عليه السلام بعد قليل", { contextCheck: false }),
      "رد T بعد قليل");

var duplicates = core.duplicateSymbols();
if (duplicates.length) {
    extra.push("  رموز مكررة في الجدول: " + duplicates.join(" / "));
} else {
    passed++;
}

var pending = core.pendingReplacements();
if (pending.length) {
    extra.push("  عبارات بلا رمز: " + pending.length);
} else {
    passed++;
}

var total = fixtures.length + 11;
var allFailures = failures.concat(extra);

console.log("");
console.log("حروف الأنوار — اختبار النواة " + core.VERSION +
            " (منقولة من السكربت " + core.SOURCE_SCRIPT_VERSION + ")");
console.log("الحالات: " + fixtures.length + " + " + (total - fixtures.length) +
            " اختباراً إضافياً");
console.log("");

if (allFailures.length) {
    console.log("إخفاقات (" + allFailures.length + "):");
    console.log(allFailures.join("\n"));
    console.log("");
    console.log("النتيجة: " + passed + " / " + total + " — لم يكتمل.");
    process.exit(1);
} else {
    console.log("النتيجة: " + passed + " / " + total + " — اجتياز كامل.");
}
