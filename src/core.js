/* =====================================================================
   حروف الأنوار — النواة المنطقية (بلا أي اعتماد على Office.js)
   منقولة من سكربت InDesign «حروف الأنوار ٣٫٢» نقلَ منطقٍ لا نقلَ شيفرة.

   القاعدة الحاكمة: بحث Word ليس GREP. فالمطابقة كلها تجري هنا على
   نصّ الفقرة بـ RegExp حقيقي، ثم تسلّم طبقةُ Word النتائجَ إلى
   paragraph.search على نطاق المطابقة وحدها.

   هذا الملف نقيّ: يعمل في المتصفح وفي Node على السواء، ولذلك
   تُختبر الحالات الـ١٠٧ عليه مباشرة بلا Word.
   ===================================================================== */

(function (root, factory) {
    "use strict";
    var api = factory();
    if (typeof module === "object" && module.exports) { module.exports = api; }
    if (root) { root.HuroofCore = api; }
})(typeof self !== "undefined" ? self :
   (typeof global !== "undefined" ? global : this), function () {
    "use strict";

    var VERSION = "2.0";
    var SOURCE_SCRIPT_VERSION = "3.2";
    var STYLE_NAME = "حروف الأنوار";
    var FONT_NAME = "AlAnwarLetters";

    /* =================================================================
       ١) البيانات — منقولة حرفياً من السكربت. لا تُغيَّر الرموز.
       ================================================================= */

    var TAIL_ALL = ["أجمعين", "جميعا"];
    var TAIL_WASALLAM = ["وسلم", "و سلم"];
    var TAIL_SHAREEF = ["الشريف"];
    var TAIL_TAALA = ["تعالى"];
    var TAIL_DHILL = ["الشريف", "الوارف"];

    /*
        حقول كل مدخل:
        find    : العبارة المعتمدة للبحث، وهي نفسها نص الاسترجاع.
        replace : حرف الرمز في خط حروف الأنوار.
        family  : عائلة الموانع التي تمنع التحويل داخل جملة عادية.
        tail    : لواحق تُبتلَع مع العبارة داخل الرمز ولا تعود.
        aliases : صيغ أخرى تتحول إلى الرمز نفسه، والاسترجاع على find.
        restore : نص استرجاع مخالف لـ find عند الحاجة (اختياري).
    */
    var REPLACEMENTS = [
        { find: "صلى الله عليه وآله", replace: "M", tail: TAIL_WASALLAM },

        { find: "عليهم صلوات الله", replace: "h", tail: TAIL_ALL,
          aliases: ["صلوات الله عليهم"] },

        { find: "صلوات الله عليه", replace: "L" },
        { find: "رضوان الله عليهم", replace: "c" },
        { find: "رضوان الله عليه", replace: "b" },

        { find: "سلام الله عليهم", replace: "f", tail: TAIL_ALL },
        { find: "سلام الله عليها", replace: "e" },
        { find: "سلام الله عليه", replace: "d" },

        { find: "عجل الله تعالى فرجه", replace: "N", family: "ajjal",
          tail: TAIL_SHAREEF, aliases: ["عجل الله فرجه"] },

        { find: "عليهما السلام", replace: "V", family: "salam" },
        { find: "عليهم السلام", replace: "W", family: "salam",
          tail: TAIL_ALL },
        { find: "عليهن السلام", replace: "X", family: "salam" },
        { find: "عليها السلام", replace: "U", family: "salam" },
        { find: "عليه السلام", replace: "T", family: "salam" },

        { find: "قدس سرهما", replace: "Q", aliases: ["قدس الله سرهما"] },
        { find: "قدس سرهم", replace: "R", tail: TAIL_ALL,
          aliases: ["قدس الله سرهم"] },
        { find: "قدس سره", replace: "P", aliases: ["قدس الله سره"] },

        { find: "حفظهم الله", replace: "K", family: "hafiz",
          tail: TAIL_ALL.concat(TAIL_TAALA) },
        { find: "حفظه الله", replace: "J", family: "hafiz",
          tail: TAIL_TAALA },

        { find: "رحمها الله", replace: "H", family: "rahima",
          tail: TAIL_TAALA },
        { find: "رحمهم الله", replace: "I", family: "rahima",
          tail: TAIL_ALL.concat(TAIL_TAALA) },
        { find: "رحمه الله", replace: "G", family: "rahima",
          tail: TAIL_TAALA },

        { find: "دام ظلهما", replace: "Z", family: "dham" },
        { find: "دام ظلهم", replace: "a", family: "dham", tail: TAIL_DHILL },
        { find: "دام ظله", replace: "Y", family: "dham", tail: TAIL_DHILL },

        { find: "عز وجل", replace: "O", family: "azza" }
    ];

    /* الكلمات التي تجعل العبارة جزءاً من جملة عادية لا رمزَ تعظيم. */
    var PARTICLES = [
        "ثم", "حتى",
        "إذا", "إن", "أن", "لو", "لما", "إذ",
        "حين", "حينما", "عندما", "كلما", "متى",
        "لعل", "لكن",
        "من", "الذي", "التي", "الذين",
        "اللاتي", "اللواتي", "اللائي",
        "لمن", "ممن", "بمن", "كمن",
        "بما", "كما", "إنما", "ربما", "مهما",
        "بينما", "حيثما", "أينما", "كيفما",
        "ريثما", "طالما", "حالما", "بعدما", "مثلما"
    ];

    var SALAM_VERBS = [
        "رد", "ردد", "ردت", "رددت",
        "ردوا", "رددوا", "رددن",
        "ترد", "يرد", "تردد", "يردد",
        "أعاد", "أعادت", "يعيد", "تعيد",
        "ألقى", "ألقت", "يلقي", "تلقي",
        "أرجع", "أرجعت", "يرجع", "ترجع",
        "أبلغ", "أبلغت", "يبلغ", "تبلغ",
        "نقل", "نقلت", "ينقل", "تنقل",
        "بعث", "بعثت", "يبعث",
        "أرسل", "أرسلت", "يرسل", "ترسل",
        "بدأ", "بدأت", "ابتدأ",
        "حيا", "حيت", "يحيي", "تحيي",
        "بلغ", "بلغت", "بلغوا", "يبلغه", "تبلغه",
        "سلم", "سلمت", "سلموا", "يسلم", "تسلم",
        "أفشى", "أفشت", "أفشوا", "يفشي", "تفشي", "افش",
        "نشر", "نشرت", "نشروا", "ينشر", "تنشر",
        "قرأ", "قرأت", "اقرأ", "اقرئ", "يقرأ", "تقرأ",
        "أهدى", "أهدت", "يهدي", "تهدي",
        "خص", "خصت", "يخص", "تخص"
    ];

    var DHAM_PARTICLES = ["ما", "منذ", "مذ", "لطالما"];

    var FAMILY_BLOCKERS = {
        salam: SALAM_VERBS,
        dham: DHAM_PARTICLES,
        azza: [],
        ajjal: [],
        hafiz: [],
        rahima: []
    };

    /* خيارات المسافة السابقة للرمز. */
    var SPACE_OPTIONS = [
        { id: "none",   label: "بدون تغيير", value: null,
          hint: "تبقى المسافة السابقة كما هي." },
        { id: "nbsp",   label: "غير منقسمة", value: "\u00A0",
          hint: "تمنع انفصال الرمز عن الاسم عند نهاية السطر." },
        { id: "narrow", label: "غير منقسمة ضيقة", value: "\u202F",
          hint: "غير منقسمة وأضيق من المسافة العادية." },
        { id: "thin",   label: "رفيعة", value: "\u2009",
          hint: "مسافة رقيقة، لكنها قد تنقسم عند نهاية السطر." }
    ];

    /* =================================================================
       ٢) لبنات التعابير النمطية
       ================================================================= */

    var MARKS = "[\u0640\u064B-\u065F\u0670\u06D6-\u06ED]*";
    var SPACES = "[ \u00A0\u202F\u2009\t]+";
    var SPACES_OPT = "[ \u00A0\u202F\u2009\t]*";
    var DASH = "[\u002D\u2010\u2011\u2013\u2014\u2212]";

    var ALEF_FORMS = "\u0627\u0623\u0625\u0622\u0671";
    var ALEF_CLASS = "[\u0627\u0623\u0625\u0622\u0671]";
    var YAA_FORMS = "\u064A\u0649";
    var YAA_CLASS = "[\u064A\u0649]";

    /*
        حدّ الكلمة. لم يُستعمل \p{L} لأنه يوجب الراية u، ولا استُعمل
        lookbehind لأن محرك Word القديم يعجز عنهما. فالحدّ الأيمن
        lookahead داخل النمط — وهو لازم هناك ليصح التراجع فتُترك
        اللاحقة إن أفسدت الحدّ: «دام ظله الشريفان» — والحدّ الأيسر
        يُفحص برمجياً على الحرف السابق، ولا يحتاج تراجعاً.

        والعلامات العربية داخل الصنف عمداً: التطويل والحركات جزء من
        الكلمة. أما «،» و«؛» و«؟» و«۔» فخارجه فهي ترقيم.
    */
    var WORD_CHAR_SRC =
        "[A-Za-z0-9_\u0621-\u065F\u0660-\u0669" +
        "\u066E-\u06D3\u06D5-\u06FF\u0750-\u077F" +
        "\uFB50-\uFDFF\uFE70-\uFEFF]";
    var RIGHT_BOUNDARY = "(?!" + WORD_CHAR_SRC + ")";
    var WORD_CHAR_RE = new RegExp("^" + WORD_CHAR_SRC + "$");

    function escapeChar(ch) {
        return "\\.^$|?*+()[]{}/".indexOf(ch) >= 0 ? "\\" + ch : ch;
    }

    function isPhraseSpace(ch) {
        return ch === " " || ch === "\u00A0" || ch === "\u202F" ||
               ch === "\u2009" || ch === "\t";
    }

    /* المسافات التي يديرها الخيار: بلا الجدولة. */
    function isManagedSpace(ch) {
        return ch === " " || ch === "\u00A0" ||
               ch === "\u202F" || ch === "\u2009";
    }

    function isSpecialSpace(ch) {
        return ch === "\u00A0" || ch === "\u202F" || ch === "\u2009";
    }

    function isMarkChar(ch) {
        var c = ch.charCodeAt(0);
        return c === 0x0640 ||
               (c >= 0x064B && c <= 0x065F) ||
               c === 0x0670 ||
               (c >= 0x06D6 && c <= 0x06ED);
    }

    function isWordChar(ch) {
        return !!ch && WORD_CHAR_RE.test(ch);
    }

    /* يطابق الحرف وصوره، والناتج دائماً بطول حرف واحد في النص. */
    function letterPattern(ch) {
        if (ALEF_FORMS.indexOf(ch) >= 0) { return ALEF_CLASS; }
        if (YAA_FORMS.indexOf(ch) >= 0) { return YAA_CLASS; }
        return escapeChar(ch);
    }

    /* كل حرف عربي يتبعه صنف الحركات، فتستوي «صلّى» و«صلى». */
    function buildPhrasePattern(phrase) {
        var out = "";
        var inSpaceRun = false;
        for (var i = 0; i < phrase.length; i++) {
            var ch = phrase.charAt(i);
            if (isPhraseSpace(ch)) {
                if (!inSpaceRun) { out += SPACES; }
                inSpaceRun = true;
                continue;
            }
            inSpaceRun = false;
            if (isMarkChar(ch)) { continue; }
            out += letterPattern(ch) + MARKS;
        }
        return out;
    }

    /*
        اللواحق تتراكم: النجمة لا علامة الاستفهام، فـ «رحمهم الله
        تعالى أجمعين» تُبتلَع لاحقتاها معاً بأي ترتيب.
    */
    function buildTailPattern(entry) {
        if (!entry.tail || !entry.tail.length) { return ""; }
        var parts = [];
        for (var i = 0; i < entry.tail.length; i++) {
            parts.push(buildPhrasePattern(entry.tail[i]));
        }
        return "(?:" + SPACES + "(?:" + parts.join("|") + "))*";
    }

    var patternCache = {};

    /*
        النمطان مرتبان: المحاط بشرطتين أولاً ثم المجرد. والترتيب لازم؛
        فلو سبق المجردُ لالتهم العبارة وحدها وترك الشرطتين يتيمتين:
        «الشيخ - P - كتب». وتُشترط الشرطتان معاً، والمفردة تُصان.
    */
    function patternsFor(entryIndex, entry, phrase, stripDashes) {
        var key = entryIndex + "|" + phrase + "|" + (stripDashes ? 1 : 0);
        if (patternCache[key]) { return patternCache[key]; }

        var core = buildPhrasePattern(phrase) +
                   buildTailPattern(entry) +
                   RIGHT_BOUNDARY;
        var list = [];

        if (stripDashes) {
            list.push({
                dashed: true,
                re: new RegExp(
                    "(" + DASH + SPACES_OPT + ")(" + core + ")(" +
                    SPACES_OPT + DASH + ")", "g")
            });
        }
        list.push({ dashed: false, re: new RegExp("(" + core + ")", "g") });

        patternCache[key] = list;
        return list;
    }

    /* =================================================================
       ٣) حكم السياق — ثلاث دوال نقيّة
       ================================================================= */

    /* تجريد الحركات والتطويل وتوحيد صور الألف والياء. */
    function normalizeWord(word) {
        var out = "";
        for (var i = 0; i < word.length; i++) {
            var ch = word.charAt(i);
            var c = ch.charCodeAt(0);

            if (c === 0x0640) { continue; }
            if (c >= 0x064B && c <= 0x065F) { continue; }
            if (c === 0x0670) { continue; }
            if (c >= 0x06D6 && c <= 0x06ED) { continue; }

            if (c === 0x0623 || c === 0x0625 ||
                c === 0x0622 || c === 0x0671) {
                ch = "\u0627";
            } else if (c === 0x0649) {
                ch = "\u064A";
            }
            out += ch;
        }
        return out;
    }

    function isArabicWordChar(ch) {
        var c = ch.charCodeAt(0);
        return c >= 0x0621 && c <= 0x06FF;
    }

    /*
        آخر كلمة في نصٍّ سابق. تُتخطى المسافات وحدها؛ فإن كان ما قبلها
        علامةَ ترقيم أو قوساً أو أولَ فقرة فلا كلمة قبلها، والعبارة
        حينئذ لقب لا فعل، فلا حجب.
    */
    function lastWordOf(before) {
        var i = before.length - 1;

        while (i >= 0 && isPhraseSpace(before.charAt(i))) { i--; }
        if (i < 0) { return ""; }
        if (!isArabicWordChar(before.charAt(i))) { return ""; }

        var end = i;
        while (i >= 0 && isArabicWordChar(before.charAt(i))) { i--; }
        return before.substring(i + 1, end + 1);
    }

    var blockerCache = {};

    function blockerSet(family) {
        var key = "@" + family;
        if (blockerCache[key]) { return blockerCache[key]; }

        var set = {};
        var i;
        for (i = 0; i < PARTICLES.length; i++) {
            set["@" + normalizeWord(PARTICLES[i])] = true;
        }
        var extra = FAMILY_BLOCKERS[family];
        if (extra) {
            for (i = 0; i < extra.length; i++) {
                set["@" + normalizeWord(extra[i])] = true;
            }
        }
        blockerCache[key] = set;
        return set;
    }

    /*
        المقارنة بالكلمة كاملةً مجرَّدةً، لا بأطراف الحروف. وتُجرَّب
        بعد نزع واو العطف أو فائه: «وردّ» و«فحين».
    */
    function isBlockingWord(word, family) {
        if (!word || !word.length) { return false; }

        var normalized = normalizeWord(word);
        if (!normalized.length) { return false; }

        var set = blockerSet(family);
        if (set["@" + normalized]) { return true; }

        var first = normalized.charAt(0);
        if ((first === "\u0648" || first === "\u0641") &&
            normalized.length > 1) {
            if (set["@" + normalized.substring(1)]) { return true; }
        }
        return false;
    }

    /* =================================================================
       ٤) قراءة الجدول
       ================================================================= */

    function symbolOf(entry) {
        if (entry.replace === undefined || entry.replace === null) { return ""; }
        return String(entry.replace);
    }

    function entryPhrases(entry) {
        var list = [entry.find];
        if (entry.aliases) {
            for (var i = 0; i < entry.aliases.length; i++) {
                if (list.indexOf(entry.aliases[i]) < 0) {
                    list.push(entry.aliases[i]);
                }
            }
        }
        return list;
    }

    /* نص الاسترجاع لا يشمل اللواحق ولا الشرطتين أبداً. */
    function restoreTextOf(entry) {
        return entry.restore ? entry.restore : entry.find;
    }

    function activeReplacements() {
        var out = [];
        for (var i = 0; i < REPLACEMENTS.length; i++) {
            if (symbolOf(REPLACEMENTS[i]).length) { out.push(REPLACEMENTS[i]); }
        }
        return out;
    }

    function pendingReplacements() {
        var out = [];
        for (var i = 0; i < REPLACEMENTS.length; i++) {
            if (!symbolOf(REPLACEMENTS[i]).length) { out.push(REPLACEMENTS[i]); }
        }
        return out;
    }

    /* الرمز الواحد لا يصلح لعبارتين، لأن الاسترجاع يقرأ الرمز وحده. */
    function duplicateSymbols() {
        var list = activeReplacements();
        var seen = {};
        var duplicates = [];
        for (var i = 0; i < list.length; i++) {
            var symbol = symbolOf(list[i]);
            var key = "@" + symbol;
            if (seen[key]) {
                duplicates.push("«" + symbol + "» بين " + seen[key] +
                                " و" + list[i].find);
            } else {
                seen[key] = list[i].find;
            }
        }
        return duplicates;
    }

    var LATIN_RE = /^[A-Za-z]$/;

    function isLatinNonSymbol(ch, map) {
        return !!ch && LATIN_RE.test(ch) && !map[ch];
    }

    function symbolIndex() {
        var map = {};
        var list = activeReplacements();
        for (var i = 0; i < list.length; i++) {
            var s = symbolOf(list[i]);
            if (!map[s]) { map[s] = list[i]; }
        }
        return map;
    }

    /* =================================================================
       ٥) تخطيط التحويل على نصّ فقرة
       ================================================================= */

    function overlaps(taken, start, end) {
        for (var i = 0; i < taken.length; i++) {
            if (start < taken[i][1] && end > taken[i][0]) { return true; }
        }
        return false;
    }

    function emptyStats() {
        return {
            matched: 0,       /* ما قُبل للتحويل */
            blocked: 0,       /* المحجوب بالسياق */
            unchecked: 0,     /* ما تعذّر فحص سياقه */
            byPhrase: {},     /* تفصيل العدد لكل عبارة */
            blockedSamples: []
        };
    }

    function countPhrase(stats, name) {
        stats.byPhrase[name] = (stats.byPhrase[name] || 0) + 1;
    }

    /*
        العائد: { ops, stats }
        كل عملية: { start, end, literal, symbol, entryIndex, phrase,
                    spaceLead, newSpace }
        و«literal» هو نصّ المطابقة حرفياً كما هو في الفقرة، ليُسلَّم
        إلى paragraph.search بلا تحوير.
    */
    function planText(text, options) {
        options = options || {};
        var spaceChar = options.spaceChar || null;
        var stripDashes = options.stripDashes !== false;
        var contextCheck = options.contextCheck !== false;

        var ops = [];
        var taken = [];
        var stats = emptyStats();

        if (!text || !text.length) { return { ops: ops, stats: stats }; }

        var entries = REPLACEMENTS;

        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            if (!symbolOf(entry).length) { continue; }

            var phrases = entryPhrases(entry);
            for (var p = 0; p < phrases.length; p++) {
                var specs = patternsFor(i, entry, phrases[p], stripDashes);

                for (var k = 0; k < specs.length; k++) {
                    var spec = specs[k];
                    var re = spec.re;
                    re.lastIndex = 0;
                    var m;

                    while ((m = re.exec(text)) !== null) {
                        var mStart = m.index;
                        var mEnd = m.index + m[0].length;
                        var coreStart = spec.dashed ?
                            mStart + m[1].length : mStart;

                        /* الحدّ الأيسر: الحرف السابق للعبارة نفسها. */
                        if (coreStart > 0 &&
                            isWordChar(text.charAt(coreStart - 1))) {
                            re.lastIndex = mStart + 1;
                            continue;
                        }

                        if (overlaps(taken, mStart, mEnd)) {
                            re.lastIndex = mStart + 1;
                            continue;
                        }

                        /* حكم السياق على الكلمة السابقة. */
                        if (contextCheck && entry.family) {
                            var word = lastWordOf(text.substring(0, coreStart));
                            if (isBlockingWord(word, entry.family)) {
                                stats.blocked++;
                                if (stats.blockedSamples.length < 12) {
                                    stats.blockedSamples.push({
                                        phrase: phrases[p],
                                        word: word
                                    });
                                }
                                re.lastIndex = mStart + 1;
                                continue;
                            }
                        }

                        var start = mStart;
                        var spaceLead = false;
                        if (spaceChar && start > 0) {
                            var prev = text.charAt(start - 1);
                            if (isManagedSpace(prev) && prev !== spaceChar &&
                                !overlaps(taken, start - 1, start)) {
                                start = start - 1;
                                spaceLead = true;
                            }
                        }

                        ops.push({
                            start: start,
                            end: mEnd,
                            literal: text.substring(start, mEnd),
                            symbol: symbolOf(entry),
                            entryIndex: i,
                            phrase: phrases[p],
                            find: entry.find,
                            spaceLead: spaceLead,
                            newSpace: spaceLead ? spaceChar : null
                        });

                        taken.push([start, mEnd]);
                        stats.matched++;
                        countPhrase(stats, entry.find);
                        re.lastIndex = mEnd;
                    }
                }
            }
        }

        ops.sort(function (a, b) { return a.start - b.start; });
        return { ops: ops, stats: stats };
    }

    /* تطبيق الخطة على سلسلة نصية — للاختبار والمعاينة الذهنية. */
    function applyOps(text, ops) {
        var out = "";
        var cursor = 0;
        for (var i = 0; i < ops.length; i++) {
            var op = ops[i];
            out += text.substring(cursor, op.start);
            out += (op.spaceLead ? op.newSpace : "") + op.symbol;
            cursor = op.end;
        }
        out += text.substring(cursor);
        return out;
    }

    function convertText(text, options) {
        var plan = planText(text, options);
        return applyOps(text, plan.ops);
    }

    /* =================================================================
       ٦) تخطيط الاسترجاع
       ================================================================= */

    /*
        كل عملية: { start, end, literal, text, symbol, entryIndex,
                    symbolIndexInText, spaceLead }
        وتُعاد المسافة الخاصة السابقة إلى مسافة عادية، لأنها من صنع
        التحويل لا من صنع المؤلف.
    */
    function planRestore(text, options) {
        options = options || {};
        var normalizeSpace = options.normalizeSpace !== false;

        var ops = [];
        if (!text || !text.length) { return { ops: ops }; }

        var map = symbolIndex();
        var seen = {};

        for (var i = 0; i < text.length; i++) {
            var ch = text.charAt(i);
            var entry = map[ch];
            if (!entry) { continue; }

            /*
                ترتيب الرمز يُحسب على كل ظهور له في الفقرة، مقبولاً كان
                أو مرفوضاً، ليطابق ما يجده بحث Word فتصح المزاوجة.
            */
            seen[ch] = (seen[ch] || 0);
            var ordinal = seen[ch];
            seen[ch] = ordinal + 1;

            /*
                الرمز لا يقع قط داخل كلمة لاتينية، فحرف «W» في كلمة
                Word ليس رمزاً. ويُستثنى الجار الذي هو رمز بنفسه،
                لئلا يسقط رمزان تجاورا. وهذا ترشيح أوّلي يوفّر بحثاً
                لا طائل تحته؛ والحكم الأخير لنمط الحرف في طبقة Word.
            */
            if (isLatinNonSymbol(text.charAt(i - 1), map) ||
                isLatinNonSymbol(text.charAt(i + 1), map)) {
                continue;
            }

            var start = i;
            var replacement = restoreTextOf(entry);
            var spaceLead = false;

            if (normalizeSpace && i > 0 && isSpecialSpace(text.charAt(i - 1))) {
                start = i - 1;
                replacement = " " + replacement;
                spaceLead = true;
            }

            ops.push({
                start: start,
                end: i + 1,
                literal: text.substring(start, i + 1),
                text: replacement,
                symbol: ch,
                symbolOrdinal: ordinal,
                find: entry.find,
                spaceLead: spaceLead
            });
        }
        return { ops: ops };
    }

    function restoreText(text, options) {
        var plan = planRestore(text, options);
        var out = "";
        var cursor = 0;
        for (var i = 0; i < plan.ops.length; i++) {
            var op = plan.ops[i];
            out += text.substring(cursor, op.start);
            out += op.text;
            cursor = op.end;
        }
        out += text.substring(cursor);
        return out;
    }

    /* =================================================================
       ٧) أدوات مشتركة لطبقة Word
       ================================================================= */

    /* ترتيب ظهور نصٍّ حرفي داخل فقرة، كما يعدّه بحث Word. */
    function literalOccurrences(text, literal) {
        var positions = [];
        if (!literal || !literal.length) { return positions; }
        var pos = 0;
        var k;
        while ((k = text.indexOf(literal, pos)) >= 0) {
            positions.push(k);
            pos = k + literal.length;
        }
        return positions;
    }

    function normalizeFontName(value) {
        if (value === undefined || value === null) { return ""; }
        return String(value).toLowerCase().replace(/[\s_\-\t]/g, "");
    }

    function isGlyphFont(value) {
        return normalizeFontName(value) === normalizeFontName(FONT_NAME);
    }

    function spaceOptionById(id) {
        for (var i = 0; i < SPACE_OPTIONS.length; i++) {
            if (SPACE_OPTIONS[i].id === id) { return SPACE_OPTIONS[i]; }
        }
        return SPACE_OPTIONS[0];
    }

    return {
        VERSION: VERSION,
        SOURCE_SCRIPT_VERSION: SOURCE_SCRIPT_VERSION,
        STYLE_NAME: STYLE_NAME,
        FONT_NAME: FONT_NAME,

        REPLACEMENTS: REPLACEMENTS,
        PARTICLES: PARTICLES,
        SALAM_VERBS: SALAM_VERBS,
        DHAM_PARTICLES: DHAM_PARTICLES,
        FAMILY_BLOCKERS: FAMILY_BLOCKERS,
        SPACE_OPTIONS: SPACE_OPTIONS,

        normalizeWord: normalizeWord,
        lastWordOf: lastWordOf,
        isBlockingWord: isBlockingWord,
        isWordChar: isWordChar,
        isSpecialSpace: isSpecialSpace,

        entryPhrases: entryPhrases,
        restoreTextOf: restoreTextOf,
        activeReplacements: activeReplacements,
        pendingReplacements: pendingReplacements,
        duplicateSymbols: duplicateSymbols,
        symbolIndex: symbolIndex,

        planText: planText,
        applyOps: applyOps,
        convertText: convertText,
        planRestore: planRestore,
        restoreText: restoreText,

        literalOccurrences: literalOccurrences,
        normalizeFontName: normalizeFontName,
        isGlyphFont: isGlyphFont,
        spaceOptionById: spaceOptionById
    };
});
