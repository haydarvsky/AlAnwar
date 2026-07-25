/* =====================================================================
   حروف الأنوار — طبقة Word والواجهة
   تعتمد على core.js في كل ما يخصّ المطابقة والحكم، ولا تُعيد منطقاً.

   قواعد لا تُخالف في هذا الملف:
   ١) لا insertText على فقرة كاملة أبداً — يمحو التنسيق الداخلي.
      الاستبدال على نطاق المطابقة وحدها.
   ٢) القراءة كلها قبل الكتابة كلها: لا sync داخل حلقة المطابقة.
   ٣) عند تكرار النصّ الحرفي في الفقرة تُزاوَج المطابقات بالفهرس،
      وإن اختلّ العدد تُترك العبارة ولا تُخمَّن.
   ===================================================================== */

(function () {
    "use strict";

    var Core = null;
    var busy = false;
    var fontReady = false;
    var fontOverride = false;      /* أذن المستخدم بالتنفيذ رغم غياب الخط */
    var apiLevel = { v13: false, v15: false, v19: false };
    var paragraphStyleNames = [];
    var defaultCharStyle = null;   /* نظير [None] في InDesign */

    /* ---------------------------------------------------------------
       أدوات الواجهة
       --------------------------------------------------------------- */

    function el(id) { return document.getElementById(id); }

    function setStatus(kind, html) {
        var box = el("status");
        box.className = kind || "";
        box.innerHTML = html || "";
    }

    function setProgress(text, percent) {
        el("progressText").textContent = text || "";
        var bar = el("progressBar");
        if (percent === undefined || percent === null) {
            el("progress").style.display = text ? "block" : "none";
            bar.style.width = "0%";
        } else {
            el("progress").style.display = "block";
            bar.style.width = Math.max(0, Math.min(100, percent)) + "%";
        }
        return new Promise(function (done) { setTimeout(done, 0); });
    }

    function setBusy(value) {
        busy = value;
        el("btnConvert").disabled = value;
        el("btnRestore").disabled = value;
        el("btnStyles").disabled = value;
    }

    function esc(text) {
        return String(text)
            .replace(/&/g, "&amp;").replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
    }

    function errorText(err) {
        if (!err) { return "خطأ غير معروف."; }
        var text = err.message || String(err);
        if (err.debugInfo && err.debugInfo.errorLocation) {
            text += " (" + err.debugInfo.errorLocation + ")";
        }
        return text;
    }

    /* ---------------------------------------------------------------
       فحص تثبيت الخط في الجهاز (لا سبيل إليه عبر Office.js)
       يُقاس عرض النصّ بالخط المطلوب ثم بخطوط النظام؛ فإن اختلف
       فالخط موجود. وغيابُه يعني طباعة حروف لاتينية مكان العبارات،
       وهو فساد صامت، فيوقَف التحويل.
       --------------------------------------------------------------- */

    function fontIsInstalled(name) {
        try {
            var probe = "MWMWmwmwILil8503";
            var canvas = document.createElement("canvas");
            var ctx = canvas.getContext("2d");
            var bases = ["monospace", "serif", "sans-serif"];
            var i, base;

            for (i = 0; i < bases.length; i++) {
                base = bases[i];
                ctx.font = "72px " + base;
                var plain = ctx.measureText(probe).width;
                ctx.font = '72px "' + name + '", ' + base;
                if (ctx.measureText(probe).width !== plain) { return true; }
            }

            if (document.fonts && document.fonts.check) {
                try {
                    if (document.fonts.check('12px "' + name + '"')) {
                        return true;
                    }
                } catch (e) {}
            }
            return false;
        } catch (e) {
            return false;
        }
    }

    function refreshFontState() {
        fontReady = fontIsInstalled(Core.FONT_NAME);
        var box = el("fontState");
        box.className = "state " + (fontReady ? "ok" : "bad");
        box.innerHTML = fontReady ?
            "الخط <b>" + esc(Core.FONT_NAME) + "</b> مثبَّت — جاهز للتحويل." :
            "الخط <b>" + esc(Core.FONT_NAME) + "</b> غير مثبَّت على هذا " +
            "الجهاز. التحويل الآن يطبع حروفاً لاتينية مكان العبارات.";
        el("fontOverrideRow").style.display = fontReady ? "none" : "flex";
        if (fontReady) { el("cbFontOverride").checked = false; }
        fontOverride = false;
    }

    /* ---------------------------------------------------------------
       الخيارات
       --------------------------------------------------------------- */

    function readOptions() {
        var scope = document.querySelector('input[name="scope"]:checked').value;
        var spaceOption = Core.spaceOptionById(el("selSpace").value);
        return {
            scope: scope,
            footnotes: el("cbFootnotes").checked,
            endnotes: el("cbEndnotes").checked,
            headers: el("cbHeaders").checked,
            shapes: el("cbShapes").checked,
            includeStyle: el("selInclude").value || "",
            excludeStyle: el("selExclude").value || "",
            spaceChar: spaceOption.value,
            stripDashes: el("cbDashes").checked,
            contextCheck: el("cbContext").checked
        };
    }

    /* ---------------------------------------------------------------
       الأنماط
       --------------------------------------------------------------- */

    var DEFAULT_CHAR_STYLE_NAMES = [
        "Default Paragraph Font",
        "خط الفقرة الافتراضي",
        "الخط الافتراضي للفقرة",
        "Absatz-Standardschriftart",
        "Police par défaut",
        "Fuente de párrafo predeterminada",
        "Carattere predefinito paragrafo",
        "Varsayılan Paragraf Yazı Tipi",
        "Standardstycketeckensnitt"
    ];

    function looksLikeDefaultCharStyle(name) {
        return /default paragraph font|الافتراضي|standardschriftart|par d[ée]faut|predeterminad|predefinit|varsay|standardstycke/i
            .test(name || "");
    }

    async function loadStyleNames(ctx) {
        var styles = ctx.document.getStyles();
        styles.load("items/nameLocal,items/type,items/builtIn");
        await ctx.sync();

        var paragraphs = [];
        var defaultChar = null;
        var i, s;

        for (i = 0; i < styles.items.length; i++) {
            s = styles.items[i];
            if (s.type === Word.StyleType.paragraph || s.type === "Paragraph") {
                paragraphs.push(s.nameLocal);
            } else if (s.type === Word.StyleType.character ||
                       s.type === "Character") {
                if (!defaultChar &&
                    (DEFAULT_CHAR_STYLE_NAMES.indexOf(s.nameLocal) >= 0 ||
                     (s.builtIn && looksLikeDefaultCharStyle(s.nameLocal)))) {
                    defaultChar = s.nameLocal;
                }
            }
        }

        paragraphs.sort(function (a, b) { return a.localeCompare(b, "ar"); });
        return { paragraphs: paragraphs, defaultChar: defaultChar };
    }

    function fillStyleSelects(names) {
        var lists = [
            { node: el("selInclude"), first: "كل أنماط الفقرات" },
            { node: el("selExclude"), first: "بدون استثناء" }
        ];

        lists.forEach(function (item) {
            var previous = item.node.value;
            item.node.innerHTML = "";
            var opt = document.createElement("option");
            opt.value = "";
            opt.textContent = item.first;
            item.node.appendChild(opt);

            names.forEach(function (name) {
                var o = document.createElement("option");
                o.value = name;
                o.textContent = name;
                item.node.appendChild(o);
            });

            if (previous && names.indexOf(previous) >= 0) {
                item.node.value = previous;
            }
        });
    }

    async function refreshStyles(silent) {
        if (!apiLevel.v15) {
            el("styleRow").style.display = "none";
            return;
        }
        try {
            await Word.run(async function (ctx) {
                var found = await loadStyleNames(ctx);
                paragraphStyleNames = found.paragraphs;
                defaultCharStyle = found.defaultChar;
                fillStyleSelects(found.paragraphs);
            });
            if (!silent) {
                setStatus("ok", "حُدِّثت قائمة أنماط الفقرات: " +
                          paragraphStyleNames.length + " نمطاً.");
            }
        } catch (err) {
            if (!silent) {
                setStatus("err", "تعذر قراءة أنماط الفقرات: " +
                          esc(errorText(err)));
            }
        }
    }

    /*
        نمط الحرف يُنشأ مرة واحدة ويُربط بالخط في كل تشغيل،
        فلو غُيِّر خطه يدوياً عاد إلى الصواب.
    */
    async function ensureGlyphStyle(ctx) {
        var styles = ctx.document.getStyles();
        var style = styles.getByNameOrNullObject(Core.STYLE_NAME);
        style.load("isNullObject");
        await ctx.sync();

        if (style.isNullObject) {
            style = ctx.document.addStyle(Core.STYLE_NAME,
                                          Word.StyleType.character);
            await ctx.sync();
        }
        style.font.name = Core.FONT_NAME;
        await ctx.sync();
        return style;
    }

    /* ---------------------------------------------------------------
       جمع الفقرات من النطاق المطلوب
       --------------------------------------------------------------- */

    async function gatherSources(ctx, opts, notes) {
        var sources = [];

        function push(label, collection) {
            sources.push({ label: label, col: collection });
        }

        if (opts.scope === "selection") {
            push("التحديد", ctx.document.getSelection().paragraphs);
            return sources;
        }

        push("المتن", ctx.document.body.paragraphs);

        if (opts.headers) {
            try {
                var sections = ctx.document.sections;
                sections.load("items");
                await ctx.sync();
                var kinds = ["Primary", "FirstPage", "EvenPages"];
                sections.items.forEach(function (section) {
                    kinds.forEach(function (kind) {
                        push("رأس/تذييل", section.getHeader(kind).paragraphs);
                        push("رأس/تذييل", section.getFooter(kind).paragraphs);
                    });
                });
            } catch (e) {
                notes.push("تعذر الوصول إلى الرؤوس والتذييلات.");
            }
        }

        if (opts.footnotes) {
            try {
                var footnotes = ctx.document.body.footnotes;
                footnotes.load("items");
                await ctx.sync();
                footnotes.items.forEach(function (note) {
                    push("حاشية", note.body.paragraphs);
                });
            } catch (e) {
                notes.push("الحواشي السفلية غير مدعومة في هذا الإصدار من Word.");
            }
        }

        if (opts.endnotes) {
            try {
                var endnotes = ctx.document.body.endnotes;
                endnotes.load("items");
                await ctx.sync();
                endnotes.items.forEach(function (note) {
                    push("تعليق ختامي", note.body.paragraphs);
                });
            } catch (e) {
                notes.push("التعليقات الختامية غير مدعومة في هذا الإصدار.");
            }
        }

        if (opts.shapes) {
            try {
                var shapes = ctx.document.body.shapes;
                shapes.load("items");
                await ctx.sync();
                shapes.items.forEach(function (shape) {
                    push("مربع نص", shape.textFrame.textRange.paragraphs);
                });
            } catch (e) {
                notes.push("مربعات النص والأشكال غير مدعومة في هذا الإصدار.");
            }
        }

        return sources;
    }

    /*
        تُحمَّل نصوص الفقرات كلها دفعةً واحدة. فإن أخفقت الدفعة —
        لاختلاف الدعم بين الحاضنات — حُمِّل كل مصدر وحده وتُخطّي
        المتعذّر منه، فلا تسقط العملية كلها بسبب مصدر واحد.
    */
    async function loadParagraphs(ctx, sources, notes) {
        sources.forEach(function (s) {
            s.col.load("items/text,items/style,items/font/name");
        });

        try {
            await ctx.sync();
        } catch (e) {
            for (var i = 0; i < sources.length; i++) {
                try {
                    sources[i].col.load("items/text,items/style,items/font/name");
                    await ctx.sync();
                } catch (inner) {
                    sources[i].failed = true;
                }
            }
            notes.push("تعذر قراءة بعض المصادر، وعُولج ما أمكن.");
        }

        var list = [];
        sources.forEach(function (s) {
            if (s.failed) { return; }
            var items;
            try { items = s.col.items; } catch (e) { return; }
            if (!items) { return; }
            items.forEach(function (p) {
                var text = "";
                try { text = p.text || ""; } catch (e) { return; }
                if (!text) { return; }
                var styleName = "";
                try { styleName = p.style || ""; } catch (e) {}
                var fontName = null;
                try { fontName = p.font ? p.font.name : null; } catch (e) {}
                list.push({
                    p: p, text: text, style: styleName,
                    font: fontName, where: s.label
                });
            });
        });
        return list;
    }

    function passesStyleFilter(item, opts, result) {
        if (opts.includeStyle && item.style !== opts.includeStyle) {
            return false;
        }
        if (opts.excludeStyle && item.style === opts.excludeStyle) {
            result.excluded++;
            return false;
        }
        return true;
    }

    /* ---------------------------------------------------------------
       التحويل
       --------------------------------------------------------------- */

    function emptyResult() {
        return {
            converted: 0, restored: 0, excluded: 0, blocked: 0,
            unchecked: 0, failed: 0, skipped: 0,
            byPhrase: {}, blockedSamples: [], notes: []
        };
    }

    async function applyConvertBatch(ctx, batch, result) {
        var jobs = [];

        batch.forEach(function (item) {
            var groups = {};
            item.ops.forEach(function (op) {
                if (!groups[op.literal]) { groups[op.literal] = []; }
                groups[op.literal].push(op);
            });

            Object.keys(groups).forEach(function (literal) {
                var ranges;
                try {
                    ranges = item.p.search(literal, {
                        matchCase: true,
                        ignorePunct: false,
                        ignoreSpace: false,
                        matchWildcards: false
                    });
                } catch (e) {
                    result.failed += groups[literal].length;
                    return;
                }
                ranges.load("items/text");
                jobs.push({
                    item: item, literal: literal,
                    ops: groups[literal], ranges: ranges
                });
            });
        });

        if (!jobs.length) { return; }

        try {
            await ctx.sync();
        } catch (e) {
            jobs.forEach(function (job) { result.failed += job.ops.length; });
            result.notes.push("تعذّر البحث في دفعة فقرات: " + errorText(e));
            return;
        }

        var applications = [];

        jobs.forEach(function (job) {
            var matches = [];
            try {
                job.ranges.items.forEach(function (r) {
                    if (r.text === job.literal) { matches.push(r); }
                });
            } catch (e) {
                result.failed += job.ops.length;
                return;
            }

            var positions = Core.literalOccurrences(job.item.text, job.literal);

            /*
                فخّ التكرار: النصّ الحرفي نفسه قد يتكرر في الفقرة،
                فتُزاوَج المطابقة بالفهرس لا بالأولى دائماً. وإن
                اختلّ العدد فالمزاوجة غير موثوقة، والخطأ الآمن أن
                تُترك العبارة لتُصلَح باليد لا أن يُخمَّن موضعها.
            */
            if (matches.length !== positions.length) {
                result.failed += job.ops.length;
                if (result.notes.length < 12) {
                    result.notes.push(
                        "تعذّر تحديد موضع «" + job.literal +
                        "» بدقة داخل فقرة، فتُركت كما هي.");
                }
                return;
            }

            job.ops.forEach(function (op) {
                var ordinal = positions.indexOf(op.start);
                if (ordinal < 0 || !matches[ordinal]) {
                    result.failed++;
                    return;
                }
                applications.push({
                    range: matches[ordinal], op: op, item: job.item
                });
            });
        });

        /* من آخر الفقرة إلى أولها، فلا تتأثر النطاقات السابقة بالطول. */
        applications.sort(function (a, b) { return b.op.start - a.op.start; });

        var queued = [];

        applications.forEach(function (a) {
            try {
                var ins = a.range.insertText(a.op.symbol,
                                             Word.InsertLocation.replace);
                ins.style = Core.STYLE_NAME;

                if (a.op.spaceLead) {
                    var space = ins.insertText(a.op.newSpace,
                                               Word.InsertLocation.before);
                    if (defaultCharStyle) {
                        space.style = defaultCharStyle;
                    } else if (a.item.font) {
                        space.font.name = a.item.font;
                    }
                }
                queued.push(a.op);
            } catch (e) {
                result.failed++;
            }
        });

        /* لا يُحتسب شيء قبل نجاح المزامنة، وفشلُ دفعةٍ لا يُسقط ما بعدها. */
        try {
            await ctx.sync();
            queued.forEach(function (op) {
                result.converted++;
                result.byPhrase[op.find] = (result.byPhrase[op.find] || 0) + 1;
            });
        } catch (e) {
            result.failed += queued.length;
            result.notes.push("تعذّر تطبيق دفعة تحويل: " + errorText(e));
        }
    }

    async function runConvert() {
        if (busy) { return; }
        var opts = readOptions();

        if (!fontReady && !el("cbFontOverride").checked) {
            setStatus("err",
                "الخط <b>" + esc(Core.FONT_NAME) + "</b> غير مثبَّت. " +
                "التحويل الآن يُنتج حروفاً لاتينية مكان العبارات. " +
                "ثبّت الخط، أو أذن بالتنفيذ من المربع أعلاه إن كنت تعلم " +
                "أن الجهاز الذي سيُطبع منه المستند فيه الخط.");
            return;
        }

        setBusy(true);
        setStatus("work", "جارٍ تجهيز التحويل…");
        var result = emptyResult();

        try {
            await Word.run(async function (ctx) {
                await setProgress("تجهيز نمط الحرف…", 2);
                await ensureGlyphStyle(ctx);

                await setProgress("جمع الفقرات…", 6);
                var sources = await gatherSources(ctx, opts, result.notes);
                var paragraphs = await loadParagraphs(ctx, sources, result.notes);

                await setProgress("مطابقة العبارات…", 14);

                var planned = [];
                paragraphs.forEach(function (item) {
                    if (!passesStyleFilter(item, opts, result)) { return; }
                    var plan = Core.planText(item.text, {
                        spaceChar: opts.spaceChar,
                        stripDashes: opts.stripDashes,
                        contextCheck: opts.contextCheck
                    });
                    result.blocked += plan.stats.blocked;
                    plan.stats.blockedSamples.forEach(function (s) {
                        if (result.blockedSamples.length < 12) {
                            result.blockedSamples.push(s);
                        }
                    });
                    if (plan.ops.length) {
                        item.ops = plan.ops;
                        planned.push(item);
                    }
                });

                if (!planned.length) { return; }

                var CHUNK = 25;
                for (var i = 0; i < planned.length; i += CHUNK) {
                    var batch = planned.slice(i, i + CHUNK);
                    await applyConvertBatch(ctx, batch, result);
                    await setProgress(
                        "التحويل: " + result.converted + " عبارة",
                        14 + Math.round(((i + CHUNK) / planned.length) * 84));
                }
            });

            renderReport("تحويل", result, opts);
        } catch (err) {
            setStatus("err", "تعذر التحويل: " + esc(errorText(err)));
        } finally {
            setProgress("");
            setBusy(false);
        }
    }

    /* ---------------------------------------------------------------
       الاسترجاع
       --------------------------------------------------------------- */

    /*
        العائد: "style" إن كان الرمز يحمل نمط الحرف، و"font" إن كان
        خط الرموز مطبَّقاً عليه تجاوزاً محلياً بلا نمط، و"" إن لم يكن
        رمزاً لنا أصلاً — كحرف لاتيني في كلمة إنجليزية.
    */
    function glyphKind(range) {
        try {
            if (range.style === Core.STYLE_NAME) { return "style"; }
        } catch (e) {}
        try {
            if (range.font && Core.isGlyphFont(range.font.name)) {
                return "font";
            }
        } catch (e2) {}
        return "";
    }

    async function applyRestoreBatch(ctx, batch, result) {
        var jobs = [];

        batch.forEach(function (item) {
            var literals = {};

            item.ops.forEach(function (op) {
                literals[op.symbol] = true;
                if (op.spaceLead) { literals[op.literal] = true; }
            });

            item.search = {};
            Object.keys(literals).forEach(function (literal) {
                var ranges;
                try {
                    ranges = item.p.search(literal, {
                        matchCase: true,
                        ignorePunct: false,
                        ignoreSpace: false,
                        matchWildcards: false
                    });
                } catch (e) {
                    return;
                }
                ranges.load("items/text,items/style,items/font/name");
                item.search[literal] = ranges;
                jobs.push(true);
            });
        });

        if (!jobs.length) { return; }

        try {
            await ctx.sync();
        } catch (e) {
            batch.forEach(function (item) { result.failed += item.ops.length; });
            result.notes.push("تعذّر البحث عن الرموز في دفعة فقرات: " +
                              errorText(e));
            return;
        }

        var applications = [];

        batch.forEach(function (item) {
            var cache = {};

            function matchesOf(literal) {
                if (cache[literal]) { return cache[literal]; }
                var ranges = item.search[literal];
                var out = null;
                if (ranges) {
                    var list = [];
                    try {
                        ranges.items.forEach(function (r) {
                            if (r.text === literal) { list.push(r); }
                        });
                    } catch (e) { list = null; }

                    if (list) {
                        var positions =
                            Core.literalOccurrences(item.text, literal);
                        if (list.length === positions.length) {
                            out = { list: list, positions: positions };
                        }
                    }
                }
                cache[literal] = out || { list: null, positions: null };
                return cache[literal];
            }

            item.ops.forEach(function (op) {
                var symbolSet = matchesOf(op.symbol);
                if (!symbolSet.list) {
                    result.failed++;
                    return;
                }

                var verify = symbolSet.list[op.symbolOrdinal];
                if (!verify) { result.failed++; return; }

                /* حرف لاتيني عادي في نصّ إنجليزي ليس رمزاً. */
                var kind = glyphKind(verify);
                if (!kind) { result.skipped++; return; }

                /*
                    op.text يحمل نصّ الاسترجاع، وتسبقه مسافة عادية إن
                    كانت قبل الرمز مسافة خاصة من صنع التحويل. فإن تعذّر
                    الوصول إلى نطاق المسافة، أُعيدت العبارة وحدها
                    وبقيت المسافة الخاصة، وهي غير ضارّة.
                */
                var target = verify;
                var text = op.spaceLead ? op.text.substring(1) : op.text;

                if (op.spaceLead) {
                    var pairSet = matchesOf(op.literal);
                    if (pairSet.list) {
                        var ordinal = pairSet.positions.indexOf(op.start);
                        if (ordinal >= 0 && pairSet.list[ordinal]) {
                            target = pairSet.list[ordinal];
                            text = op.text;
                        }
                    }
                }

                applications.push({ item: item, op: op, kind: kind,
                                    range: target, text: text });
            });
        });

        applications.sort(function (a, b) { return b.op.start - a.op.start; });

        var queued = [];

        applications.forEach(function (a) {
            try {
                var ins = a.range.insertText(a.text,
                                             Word.InsertLocation.replace);
                /*
                    الترتيب مقصود كما في السكربت: نمط الحرف الافتراضي
                    أولاً — نظير [None] — ثم مسح تجاوز الخط إن بقي.
                    وعكسه يُبقي خطّ الرموز على العبارة المسترجَعة.
                */
                var cleared = false;
                if (defaultCharStyle) {
                    ins.style = defaultCharStyle;
                    cleared = true;
                }

                /*
                    شبكة أمان السكربت نفسها: لو كان خط الرموز تجاوزاً
                    محلياً لا نمطاً، لم يكفِ مسحُ النمط، فيُردّ الخط
                    إلى خط الفقرة. ولا يُفعل هذا في الحالة السليمة
                    لئلا يُزرع تجاوز حيث لم يكن.
                */
                if (a.kind === "font" || !cleared) {
                    if (a.item.font) {
                        ins.font.name = a.item.font;
                        cleared = true;
                    }
                }

                if (!cleared && result.notes.length < 12) {
                    result.notes.push(
                        "تعذّر مسح نمط الرموز عن عبارة مسترجَعة؛ " +
                        "راجعها يدوياً.");
                }
                queued.push(a.op);
            } catch (e) {
                result.failed++;
            }
        });

        try {
            await ctx.sync();
            queued.forEach(function (op) {
                result.restored++;
                result.byPhrase[op.find] = (result.byPhrase[op.find] || 0) + 1;
            });
        } catch (e) {
            result.failed += queued.length;
            result.notes.push("تعذّر تطبيق دفعة استرجاع: " + errorText(e));
        }
    }

    async function runRestore() {
        if (busy) { return; }
        var opts = readOptions();

        setBusy(true);
        setStatus("work", "جارٍ تجهيز الاسترجاع…");
        var result = emptyResult();

        try {
            await Word.run(async function (ctx) {
                await setProgress("جمع الفقرات…", 6);
                var sources = await gatherSources(ctx, opts, result.notes);
                var paragraphs = await loadParagraphs(ctx, sources, result.notes);

                await setProgress("البحث عن الرموز…", 14);

                var planned = [];
                paragraphs.forEach(function (item) {
                    if (!passesStyleFilter(item, opts, result)) { return; }
                    var plan = Core.planRestore(item.text, {
                        normalizeSpace: true
                    });
                    if (plan.ops.length) {
                        item.ops = plan.ops;
                        planned.push(item);
                    }
                });

                if (!planned.length) { return; }

                var CHUNK = 25;
                for (var i = 0; i < planned.length; i += CHUNK) {
                    var batch = planned.slice(i, i + CHUNK);
                    await applyRestoreBatch(ctx, batch, result);
                    await setProgress(
                        "الاسترجاع: " + result.restored + " عبارة",
                        14 + Math.round(((i + CHUNK) / planned.length) * 84));
                }
            });

            renderReport("استرجاع", result, opts);
        } catch (err) {
            setStatus("err", "تعذر الاسترجاع: " + esc(errorText(err)));
        } finally {
            setProgress("");
            setBusy(false);
        }
    }

    /* ---------------------------------------------------------------
       التقرير
       --------------------------------------------------------------- */

    function renderReport(kind, result, opts) {
        var isConvert = kind === "تحويل";
        var total = isConvert ? result.converted : result.restored;
        var rows = [];

        function row(label, value, tip) {
            if (!value) { return; }
            rows.push('<div class="rrow"><span>' + esc(label) +
                      (tip ? ' <i>' + esc(tip) + '</i>' : "") +
                      '</span><b>' + value + '</b></div>');
        }

        row(isConvert ? "المحوَّل" : "المسترجَع", total);
        row("المستثنى بنمط الفقرة", result.excluded);
        row("المحجوب بالسياق", result.blocked, "عبارة داخل جملة عادية");
        row("تعذّر فحص سياقه", result.unchecked);
        row("رموز بلا نمط حروف الأنوار", result.skipped, "لم تُمَسّ");
        row("المتعذّر", result.failed);

        var detail = "";
        var names = Object.keys(result.byPhrase);
        if (names.length) {
            names.sort(function (a, b) {
                return result.byPhrase[b] - result.byPhrase[a];
            });
            detail += '<details class="detail"><summary>تفصيل العبارات (' +
                      names.length + ')</summary><div class="dlist">';
            names.forEach(function (n) {
                detail += '<div class="rrow"><span>' + esc(n) +
                          '</span><b>' + result.byPhrase[n] + '</b></div>';
            });
            detail += "</div></details>";
        }

        if (result.blockedSamples.length) {
            detail += '<details class="detail"><summary>أمثلة المحجوب ' +
                      'بالسياق</summary><div class="dlist">';
            result.blockedSamples.forEach(function (s) {
                detail += '<div class="rrow"><span>' + esc(s.phrase) +
                          '</span><b>بعد «' + esc(s.word) + '»</b></div>';
            });
            detail += "</div></details>";
        }

        if (result.notes.length) {
            detail += '<details class="detail"><summary>ملاحظات (' +
                      result.notes.length + ')</summary><div class="dlist">';
            result.notes.forEach(function (n) {
                detail += '<div class="note">' + esc(n) + "</div>";
            });
            detail += "</div></details>";
        }

        var head;
        if (total) {
            head = "تم ال" + (isConvert ? "تحويل" : "استرجاع") +
                   " — " + total + " عبارة.";
        } else if (isConvert) {
            head = "لم يُعثر على عبارات مطابقة في النطاق المحدد.";
        } else {
            head = "لا توجد رموز تحمل نمط «" + Core.STYLE_NAME +
                   "» في النطاق المحدد.";
        }

        var tail = "";
        if (!isConvert) {
            tail = '<div class="note">الاسترجاع يعيد العبارة المعتمدة ' +
                   'وحدها؛ فلا تعود اللواحق «وسلم» و«الشريف» و«أجمعين» ' +
                   'و«جميعاً» ولا الشرطتان المحيطتان، لأنها دخلت في الرمز.' +
                   "</div>";
        }

        setStatus(total ? "ok" : "warn",
                  "<b>" + esc(head) + "</b>" + rows.join("") + detail + tail);
    }

    /* ---------------------------------------------------------------
       جدول العبارات
       --------------------------------------------------------------- */

    function buildPhraseTable() {
        var body = el("phraseTable");
        var list = Core.activeReplacements();
        var html = "";

        list.forEach(function (entry) {
            var extras = [];
            if (entry.tail && entry.tail.length) {
                extras.push("لواحق: " + entry.tail.join("، "));
            }
            if (entry.aliases && entry.aliases.length) {
                extras.push("بدائل: " + entry.aliases.join("، "));
            }
            if (entry.family) { extras.push("محترزة بالسياق"); }

            html += '<tr><td class="glyph">' + esc(entry.replace) +
                    '</td><td><div>' + esc(entry.find) + "</div>" +
                    (extras.length ?
                        '<div class="tiny">' + esc(extras.join(" • ")) +
                        "</div>" : "") +
                    "</td></tr>";
        });

        body.innerHTML = html;
        el("phraseCount").textContent = list.length;

        var duplicates = Core.duplicateSymbols();
        if (duplicates.length) {
            setStatus("err", "تنبيه: رمز واحد مستعمل لأكثر من عبارة، " +
                      "والاسترجاع سيعيد الأولى منها فقط:<br>" +
                      esc(duplicates.join(" — ")));
        }
    }

    /* ---------------------------------------------------------------
       التهيئة
       --------------------------------------------------------------- */

    function wire() {
        el("btnConvert").onclick = runConvert;
        el("btnRestore").onclick = runRestore;
        el("btnStyles").onclick = function () { refreshStyles(false); };

        el("selSpace").onchange = function () {
            el("spaceHint").textContent =
                Core.spaceOptionById(this.value).hint;
        };

        el("cbFontOverride").onchange = function () {
            fontOverride = this.checked;
        };

        var space = el("selSpace");
        Core.SPACE_OPTIONS.forEach(function (o) {
            var opt = document.createElement("option");
            opt.value = o.id;
            opt.textContent = o.label;
            space.appendChild(opt);
        });
        el("spaceHint").textContent = Core.SPACE_OPTIONS[0].hint;

        document.querySelectorAll(".foldable").forEach(function (node) {
            node.querySelector(".foldhead").onclick = function () {
                node.classList.toggle("open");
            };
        });
    }

    /*
        معاينة خارج Word: لو فُتحت الصفحة في متصفح عادي لم تُهيَّأ
        Office.js أبداً، فتبقى اللوحة صامتة. فتُبنى الواجهة بعد مهلة
        ويُقال للمستخدم إنها معاينة، ويُفحص الخط — وهو فحص مفيد وحده.
    */
    var started = false;

    function startStandalone() {
        if (started) { return; }
        started = true;
        Core = window.HuroofCore;
        el("version").textContent =
            "الإصدار " + Core.VERSION + " — معاينة خارج Word";
        wire();
        buildPhraseTable();
        refreshFontState();
        el("btnConvert").disabled = true;
        el("btnRestore").disabled = true;
        el("btnStyles").disabled = true;
        el("styleRow").style.opacity = ".6";
        setStatus("warn", "<b>معاينة فقط.</b> هذه الصفحة تعمل داخل " +
                  "Microsoft Word كجزء مهام؛ وفحص تثبيت الخط أعلاه " +
                  "صحيح على كل حال.");
    }

    setTimeout(function () {
        if (!started) { startStandalone(); }
    }, 3500);

    if (typeof Office === "undefined" || !Office.onReady) {
        startStandalone();
        return;
    }

    Office.onReady(function (info) {
        if (started) { return; }
        started = true;
        Core = window.HuroofCore;
        el("version").textContent =
            "الإصدار " + Core.VERSION + " — منقول من السكربت " +
            Core.SOURCE_SCRIPT_VERSION;

        if (info.host !== Office.HostType.Word) {
            started = false;
            startStandalone();
            return;
        }

        try {
            apiLevel.v13 =
                Office.context.requirements.isSetSupported("WordApi", "1.3");
            apiLevel.v15 =
                Office.context.requirements.isSetSupported("WordApi", "1.5");
            apiLevel.v19 =
                Office.context.requirements.isSetSupported("WordApi", "1.9");
        } catch (e) {}

        wire();
        buildPhraseTable();
        refreshFontState();

        if (!apiLevel.v15) {
            el("btnConvert").disabled = true;
            el("btnRestore").disabled = true;
            setStatus("err",
                "هذا الإصدار من Word لا يدعم WordApi 1.5، وهو لازم " +
                "لإنشاء نمط الحرف وللحواشي. حدّث Word أو استعمل " +
                "Word للويب.");
            return;
        }

        if (!apiLevel.v19) {
            el("cbShapes").disabled = true;
            el("cbShapes").checked = false;
            el("shapesHint").textContent =
                "مربعات النص تحتاج Word أحدث (WordApi 1.9).";
        }

        refreshStyles(true);
    });
})();
