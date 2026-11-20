var allData = [], lyData = [];
var budgetDataGlobal = null;
var lyAvailable = false; // 昨年の行を出せるか（前年の期間が CRM のデータの開始日以降のときだけ true）
var staffStatsMaster = {}, storeStatsMaster = {}, storeGroupMap = {}, staffStoreMap = {};
var personMap = {};

// CRM（Services）のデータが入っている最初の日。前年の期間がこれより前なら昨年の行はすべて「-」
var CRM_DATA_START = "2025-12-01";

// グループと店舗の並び（列・フィルタの順）。ここに無い店舗はすべて「その他」
var OTHER_GROUP = "その他";
var GROUP_ORDER = ["京奈", "大阪", "滋三", "兵四", "旧Dj", "Gf", OTHER_GROUP];
var GROUP_STORES = {
    "京奈": ["久御山店", "171店", "精華店", "西大和店", "橿原店"],
    "大阪": ["高槻店", "八尾店", "堺大泉緑地前店", "松原天美店", "貝塚店"],
    "滋三": ["栗東店", "大津店", "彦根店", "津店", "松阪店", "鯖江店"],
    "兵四": ["神戸店", "北久米店", "久米窪田店", "高知高須店", "尼崎店"],
    "旧Dj": ["加古川店", "長田店", "福知山店", "outlet店", "舞鶴店", "熊本インター店"],
    "Gf": ["金沢店", "野々市店"]
};
var storeToGroup = {}, STORE_ORDER = [];
GROUP_ORDER.forEach(function(g) { (GROUP_STORES[g] || []).forEach(function(st) { storeToGroup[st] = g; STORE_ORDER.push(st); }); });
function groupOf(st) { return storeToGroup[st] || OTHER_GROUP; }

// COQL で読む項目。担当者名は lookup の先（Masters の Name）を一緒に取る
var SELECT_FIELDS = "ClosingDay, VisitedDateTime, nousyayoteibi, SyaryouCategory, FOrR, Seated, Option1, ServiceStore, ServicePerson, ServicePerson.Name, cancel, HanbaiCategory, Option2, Option3, Option15, Option4, Option5, Option16, Option7, Option8, Option9, Option10, Option6, BackCamera, Option17, TradeinCar, PaymentCategory, arari16, arari17, Option14, arari21, arari22, arari23, arari24, arari25, Name, ManagementNum, Deal_Name";

// 集計の期間（renderAll の中で使う）
var curRange = null;

if (typeof ZOHO !== "undefined" && ZOHO.embeddedApp) {
    ZOHO.embeddedApp.init().then(function() {
        initMonthSelector();
        setInitialDates();
        detectMyStore(); // 待たない（取れたら、表ができた後に1回だけ当てる）
        fetchByCOQL();
    });
}

// ===== サインインしている人の店舗（店舗用のアカウントはユーザー名＝店舗名）。店舗別・担当者別・日別・予算入力の最初の絞り込みに使う =====
var myUserNames = null;   // 照合用の名前（空白を除いたもの）。null = まだ取れていない、[] = 取れなかった
var myStore = null;       // 決まった店舗（表示名）。決まらなければ null（今どおり全店）
var myStoreApplied = false;
var biMyStoreUsed = false; // 予算入力で自分の店舗を最初の値に使ったか（1回だけ）
async function detectMyStore() {
    var names = [];
    try {
        var res = await ZOHO.CRM.CONFIG.getCurrentUser();
        var u = res && Array.isArray(res.users) ? res.users[0] : (res && res.full_name !== undefined ? res : null);
        if (u) [u.full_name, u.last_name, (u.first_name || "") + (u.last_name || ""), (u.last_name || "") + (u.first_name || "")].forEach(function(v) { var n = normName(v); if (n && names.indexOf(n) === -1) names.push(n); });
    } catch (e) { console.error("ユーザーの取得エラー（全店で表示します）:", e); }
    myUserNames = names;
    applyMyStoreOnce();
}
// 表ができていて、名前が取れていれば1回だけ当てる。そのあとは人の選択を保つ（月を変えても戻さない）
function applyMyStoreOnce() {
    if (myStoreApplied || myUserNames === null || loadFailed || !curRange || !Object.keys(storeStatsMaster).length) return;
    myStoreApplied = true;
    var cands = STORE_ORDER.concat(Object.keys(storeStatsMaster));
    for (var i = 0; i < cands.length && !myStore; i++) { if (cands[i] !== "未所属" && myUserNames.indexOf(normName(cands[i])) !== -1) myStore = cands[i]; }
    if (!myStore) return;
    function pick(id, v) { var sel = document.getElementById(id); if (sel && Array.prototype.some.call(sel.options, function(o) { return o.value === v; })) { sel.value = v; return true; } return false; }
    pick("group-selector", groupOf(myStore)); filterStoreByGroup();
    pick("store-selector", myStore); filterStaffByStore();
    pick("daily-filter", "s:" + myStore); renderDaily();
    // 予算入力を先に開いていて、店舗がまだ空なら、ここで自分の店舗を入れて読む（人が選んでいれば変えない）
    var bp = document.getElementById("budget-page"), ss = document.getElementById("bi-store");
    if (bp && bp.classList.contains("active") && ss && !ss.value && !biMyStoreUsed) {
        biFillSelectors();
        if (Array.prototype.some.call(ss.options, function(o) { return o.value === myStore; })) { ss.value = myStore; biMyStoreUsed = true; biLoad(); }
    }
}

function initMonthSelector() { var sel = document.getElementById('month-selector'); var now = new Date(); var curMonth = now.getFullYear() + "-" + ("0" + (now.getMonth() + 1)).slice(-2); for(var y=2025; y<=now.getFullYear() + 1; y++) { for(var m=1; m<=12; m++) { var v = y + "-" + (m < 10 ? "0" + m : m); var opt = document.createElement('option'); opt.value = v; opt.text = y + "/" + m; if(v === curMonth) opt.selected = true; sel.add(opt); } } }
function setInitialDates() { var now = new Date(); var y = now.getFullYear(), m = now.getMonth() + 1, last = new Date(y, m, 0).getDate(), mStr = ("0" + m).slice(-2); document.getElementById('start-date').value = y + "-" + mStr + "-01"; document.getElementById('end-date').value = y + "-" + mStr + "-" + last; }
function syncMonthToCalendar() { var v = document.getElementById('month-selector').value, p = v.split("-").map(Number), l = new Date(p[0], p[1], 0).getDate(); document.getElementById('start-date').value = v + "-01"; document.getElementById('end-date').value = v + "-" + l; fetchByCOQL(); }

// --- 日付の小道具（すべて "YYYY-MM-DD" の文字列で扱う） ---
function pad2(n) { return ("0" + n).slice(-2); }
function getToday() { var d = new Date(); return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }
// "2026/04/01" "2026-4-1" "2026-04-01 00:00:00" などを "2026-04-01" にそろえる
function normDate(v) { var p = String(v || "").trim().split(/[\/\-\sT]/); if (p.length < 3 || !p[0]) return ""; return p[0] + "-" + pad2(parseInt(p[1], 10)) + "-" + pad2(parseInt(p[2], 10)); }
function minStr(a, b) { return a < b ? a : b; }
function maxStr(a, b) { return a > b ? a : b; }
function lastDayOfMonth(ym) { var p = ym.split("-").map(Number); return ym + "-" + pad2(new Date(p[0], p[1], 0).getDate()); }
// 1年前の同じ日（2/29 は 2/28）
function shiftYear(ds, dy) { var p = ds.split("-").map(Number), y = p[0] + dy, m = p[1], d = p[2]; var last = new Date(y, m, 0).getDate(); if (d > last) d = last; return y + "-" + pad2(m) + "-" + pad2(d); }

function getPeriod() {
    var start = document.getElementById('start-date').value, end = document.getElementById('end-date').value;
    var month = document.getElementById('month-selector').value || (start ? start.slice(0, 7) : "");
    var today = getToday();
    var lyMonth = month ? shiftYear(month + "-01", -1).slice(0, 7) : "";
    var lyMonthStart = lyMonth ? lyMonth + "-01" : "", lyMonthEnd = lyMonth ? lastDayOfMonth(lyMonth) : "";
    var lyStart = start ? shiftYear(start, -1) : "", lyCutEnd = (start && end) ? shiftYear(minStr(end, today), -1) : "";
    return {
        start: start, end: end, month: month, today: today,
        lyMonthStart: lyMonthStart, lyMonthEnd: lyMonthEnd, lyStart: lyStart, lyCutEnd: lyCutEnd,
        lyFetchStart: minStr(lyMonthStart, lyStart), lyFetchEnd: maxStr(lyMonthEnd, lyCutEnd)
    };
}

// 読み込みの呼び出し番号。月・期間を続けて切り替えたとき、最新の呼び出しだけが描画する
var fetchSeq = 0;
var loadFailed = false; // 直近の読み込みが失敗したら true（フィルタを変えても表を出さない）
function isStale(seq) { return seq !== fetchSeq; }

// COQL を全ページ読む。失敗したら例外を投げる（握りつぶさない）。新しい呼び出しが始まったら null を返して打ち切る
// ウィジェットの COQL は1回 200 件まで（2026-11-19 実測。2000 は通らない）。COQL_PAR 回分を同時に読む
// ページの並びがずれないよう id 順で読み、id で重複を除く。order by が通らなければ外して読み直す
var COQL_PAGE = 200, COQL_PAR = 4, COQL_ORDER = " order by id";
// COQL の 0件（{"status":204,"statusText":"nocontent"}。文字列で来ることもある）
function isCoqlNoContent(x) {
    if (!x) return false;
    if (typeof x === "string") { try { x = JSON.parse(x); } catch (e) { return /nocontent/i.test(x) || /"status"\s*:\s*204/.test(x); } }
    if (typeof x !== "object") return false;
    return x.status === 204 || x.status === "204" || String(x.statusText || "").toLowerCase() === "nocontent" || String(x.code || "").toUpperCase() === "NO_CONTENT";
}
// エラーを人が読める1つの文に（message・code が無ければ中身を JSON で）
function errStr(e) {
    if (e === null || e === undefined) return "不明なエラー";
    if (typeof e === "string") return e;
    var parts = [];
    if (e.code) parts.push(e.code);
    if (e.message) parts.push(e.message);
    if (!parts.length && e.data && e.data[0]) { if (e.data[0].code) parts.push(e.data[0].code); if (e.data[0].message) parts.push(e.data[0].message); }
    if (parts.length) return parts.join(" / ");
    try { return JSON.stringify(e); } catch (x) { return String(e); }
}
// module / fields を省くと接客（Services）を SELECT_FIELDS で読む
async function coqlPage(where, offset, module, fields) {
    for (var attempt = 0; attempt < 2; attempt++) {
        var res;
        try { res = await ZOHO.CRM.API.coql({ "select_query": "select " + (fields || SELECT_FIELDS) + " from " + (module || "Services") + " where " + where + COQL_ORDER + " limit " + offset + ", " + COQL_PAGE }); }
        catch (e) {
            if (isCoqlNoContent(e)) return { rows: [], more: false }; // 0件のとき SDK が 204 の形で返す（失敗の側に来ることもある）
            if (COQL_ORDER && attempt === 0) { COQL_ORDER = ""; continue; }
            throw e;
        }
        if (res && Array.isArray(res.data)) return { rows: res.data, more: !!(res.info && res.info.more_records) };
        if (isCoqlNoContent(res)) return { rows: [], more: false };
        var bad = res && typeof res === "object" && (res.code || res.status === "error" || res.message);
        if (!bad) return { rows: [], more: false }; // 0件
        if (COQL_ORDER && attempt === 0) { COQL_ORDER = ""; continue; }
        throw new Error(res.code || res.message || "COQL error");
    }
}
async function fetchServices(where, seq, module, fields) {
    var out = [], seen = {}, offset = 0;
    while (true) {
        var offs = [];
        for (var i = 0; i < COQL_PAR; i++) offs.push(offset + i * COQL_PAGE);
        var pages = await Promise.all(offs.map(function(o) { return coqlPage(where, o, module, fields); }));
        if (seq !== undefined && seq !== null && isStale(seq)) return null;
        var more = false;
        for (var j = 0; j < pages.length; j++) {
            pages[j].rows.forEach(function(r) { var k = r.id || JSON.stringify(r); if (!seen[k]) { seen[k] = true; out.push(r); } });
            more = pages[j].more;
            if (!more) break;
        }
        if (!more) break;
        offset += COQL_PAR * COQL_PAGE;
    }
    return out;
}

async function fetchByCOQL() {
    var loadingEl = document.getElementById('loading');
    var p = getPeriod();
    // 開始日・終了日が空なら何もしない（番号を増やさないので、読み込み中の前の呼び出しは捨てない）
    if(!p.start || !p.end) { loadingEl.style.display = 'none'; return; }

    var seq = ++fetchSeq;
    loadingEl.style.display = 'block';
    loadingEl.style.textAlign = 'left';
    loadingEl.innerHTML = "<b>【データ読み込み状況】</b><br>▶ 実績データを取得中...";

    // 成約日・来店日・納車日のどれかが期間内のもの
    var where = "(((ClosingDay between '" + p.start + "' and '" + p.end + "') or (VisitedDateTime between '" + p.start + "T00:00:00+09:00' and '" + p.end + "T23:59:59+09:00')) or (nousyayoteibi between '" + p.start + "T00:00:00+09:00' and '" + p.end + "T23:59:59+09:00'))";
    // 読み終えるまでは手元の変数に持ち、最新の呼び出しのときだけ allData / lyData に入れる
    var cur = [], ly = [], lyOk = false, budget = null;
    // 今年分・昨年分・予算は互いに関係ないので同時に読む
    var budgetP = fetchSalesBudgets(p.month);
    try {
        // 昨年分は前年の期間が CRM のデータの開始日以降のときだけ読む
        var wantLy = !!(p.lyFetchStart && p.lyFetchStart >= CRM_DATA_START);
        var lyP = Promise.resolve([]);
        if (wantLy) {
            loadingEl.innerHTML += "<br>▶ 昨年の実績を取得中...";
            var lyWhere = "((ClosingDay between '" + p.lyFetchStart + "' and '" + p.lyFetchEnd + "') or (VisitedDateTime between '" + p.lyFetchStart + "T00:00:00+09:00' and '" + p.lyFetchEnd + "T23:59:59+09:00'))";
            lyP = fetchServices(lyWhere, seq);
        }
        var got = await Promise.all([fetchServices(where, seq), lyP]);
        cur = got[0]; ly = got[1];
        if (cur === null || ly === null || isStale(seq)) return;
        lyOk = wantLy;
    } catch (e) {
        if (isStale(seq)) return;
        console.error("COQL 取得エラー:", e);
        // 前の月の集計を残さない（フィルタを変えても前の表が出ないように）
        allData = []; lyData = []; lyAvailable = false; budgetDataGlobal = null;
        storeStatsMaster = {}; staffStatsMaster = {}; storeGroupMap = {}; staffStoreMap = {}; loadFailed = true;
        loadingEl.innerHTML += "<br><span style='color:#c00000;'>▶ 取得に失敗しました（" + String((e && e.message) || e).replace(/[<>&]/g, "") + "）。期間を選び直すか、画面を開き直してください。</span>";
        updateSelector('group-selector', [], '全グループ表示'); updateSelector('store-selector', [], '全店舗表示');
        ["group-table-container", "store-table-container", "staff-table-container", "daily-table-container"].forEach(function(id) { var el = document.getElementById(id); if (el) el.innerHTML = ""; });
        return;
    }

    var both = await Promise.all([resolveMastersNames(cur.concat(ly)), budgetP]);
    budget = both[1];
    if (isStale(seq)) return;

    allData = cur; lyData = ly; lyAvailable = lyOk; budgetDataGlobal = budget;
    renderAll(seq);
}

// 担当者名は COQL の ServicePerson.Name から。取れなかった id だけ 1件ずつ読む（予備）
async function resolveMastersNames(all) {
    var loadingEl = document.getElementById('loading');
    all.forEach(function(r) { var id = r.ServicePerson && r.ServicePerson.id, nm = r["ServicePerson.Name"]; if (id && nm) personMap[id] = nm; });
    var ids = [...new Set(all.map(r => r.ServicePerson ? r.ServicePerson.id : null).filter(id => id && !personMap[id]))];
    if (ids.length === 0) return;
    loadingEl.innerHTML += "<br>▶ 担当者名を照合中...";
    await Promise.all(ids.map(async (id) => {
        try { var res = await ZOHO.CRM.API.getRecord({ Entity: "Masters", RecordID: id }); if (res.data && res.data.length > 0) personMap[id] = res.data[0].Name || "名称未設定"; else personMap[id] = "ID:" + id; } catch (e) { personMap[id] = "ID:" + id; }
    }));
}

// ===== 予算・目標（CRM のモジュール「販売予算」SalesBudgets。1か月×店舗×担当者で1件、担当者が空なら店舗全体の行） =====
var BUDGET_FIELDS = "id, Name, BudgetMonth, BudgetStore, BudgetStaff, BudgetStaff.Name, BudgetKei, BudgetFu, TargetKei, TargetFu, DeliveryBudget, GrossProfitBudget";
function coqlStr(v) { return String(v === null || v === undefined ? "" : v).replace(/\\/g, "\\\\").replace(/'/g, "\\'"); }
// 同じモジュールに3種類: 月予算（BudgetType が空か「月予算」）/ 曜日の重み（1件＝1店舗、BudgetMonth＝適用開始月）/ 日別予算（1件＝1店舗×1日）
var BT_MONTH = "月予算", BT_WEIGHT = "曜日の重み", BT_DAILY = "日別予算";
var WEIGHT_FIELDS_LIST = ["WeightMon", "WeightTue", "WeightWed", "WeightThu", "WeightFri", "WeightSat", "WeightSun", "WeightHoliday"];
var WEIGHT_LABELS = ["月", "火", "水", "木", "金", "土", "日", "祝日"];
var WEIGHT_SELECT = "id, Name, BudgetType, BudgetStore, BudgetMonth, " + WEIGHT_FIELDS_LIST.join(", ");
var DAILY_SELECT = "id, Name, BudgetType, BudgetStore, BudgetMonth, TargetDate, DailyBudget";
function isMonthBudgetRow(b) { var t = b && b.BudgetType; return t === null || t === undefined || t === "" || t === "-None-" || t === BT_MONTH; }
// その月（"YYYY-MM"）の月予算の行。店舗を渡せばその店舗だけ。読めなければ例外
// 種別の条件が通らなければ、種別なしで読んで手元で月予算だけに分ける
// COQL の where は、条件を3つ以上つなぐときは2つずつ括弧でくくる（((A and B) and C)。そうしないと SYNTAX_ERROR）
// 渡す条件のうち or を含むものは、呼ぶ側で括弧に入れて1項にしておく
function coqlAnd(conds) {
    var list = conds.filter(function(c) { return !!c; });
    if (list.length === 0) return "";
    var w = list[0];
    for (var i = 1; i < list.length; i++) w = "(" + w + " and " + list[i] + ")";
    return w;
}
async function querySalesBudgets(month, store) {
    var cMonth = "BudgetMonth = '" + coqlStr(month) + "-01'", cStore = store ? "BudgetStore = '" + coqlStr(store) + "'" : null;
    try { return (await fetchServices(coqlAnd([cMonth, cStore, "(BudgetType is null or BudgetType = '" + BT_MONTH + "')"]), null, "SalesBudgets", BUDGET_FIELDS + ", BudgetType")).filter(isMonthBudgetRow); }
    catch (e) {
        console.error("月予算の種別つきの読み込みが通らないので、全件を読んで分けます:", e);
        return (await fetchServices(coqlAnd([cMonth, cStore]), null, "SalesBudgets", BUDGET_FIELDS + ", BudgetType")).filter(isMonthBudgetRow);
    }
}
// 曜日の重みの行（適用開始月が month 以前のもの。店舗を渡せばその店舗だけ）
async function queryWeightRows(month, store) {
    var where = coqlAnd(["BudgetType = '" + BT_WEIGHT + "'", "BudgetMonth <= '" + coqlStr(month) + "-01'", store ? "BudgetStore = '" + coqlStr(store) + "'" : null]);
    return await fetchServices(where, null, "SalesBudgets", WEIGHT_SELECT);
}
// 日別予算の行（その月。店舗を渡せばその店舗だけ）
async function queryDailyRows(month, store) {
    var where = coqlAnd(["BudgetType = '" + BT_DAILY + "'", "BudgetMonth = '" + coqlStr(month) + "-01'", store ? "BudgetStore = '" + coqlStr(store) + "'" : null]);
    return await fetchServices(where, null, "SalesBudgets", DAILY_SELECT);
}
// 表の読み込み用。月予算が読めなければ null（予算の行は「未登録」）。重み・日別が読めなければ空（共通の重みで割る）
async function fetchSalesBudgets(month) {
    if (!month) return null;
    try {
        var got = await Promise.all([
            querySalesBudgets(month, null),
            queryWeightRows(month, null).catch(function(e) { console.error("曜日の重みの取得エラー（共通の重みで割ります）:", e); return []; }),
            queryDailyRows(month, null).catch(function(e) { console.error("日別予算の取得エラー（手入力なしで割ります）:", e); return []; })
        ]);
        return { month: month, rows: got[0], weights: got[1], daily: got[2] };
    }
    catch (e) { console.error("予算（販売予算）の取得エラー:", e); return null; }
}

// ===== 日割り（店舗ごと） =====
// その店舗でその月に効く重みの行（BudgetMonth ≤ その月 の中で一番新しい行）。無ければ null（＝全店共通の重み）
function budWeightRowFor(weights, store, month) {
    var best = null, m1 = month + "-01";
    (weights || []).forEach(function(w) {
        if ((w.BudgetStore || "未所属") !== (store || "未所属")) return;
        var bm = String(w.BudgetMonth || "").slice(0, 10);
        if (!bm || bm > m1) return;
        if (!best || bm > String(best.BudgetMonth).slice(0, 10) || (bm === String(best.BudgetMonth).slice(0, 10) && String(w.id) > String(best.id))) best = w;
    });
    return best;
}
// 重みの行（無ければ共通）でのその日の重み。行の空の曜日は 0
function budDayWeightBy(ds, wrow) {
    if (!wrow) return budDayWeight(ds);
    var p = ds.split("-").map(Number), dow = new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
    var f = JP_HOLIDAYS.indexOf(ds) !== -1 ? "WeightHoliday" : ["WeightSun", "WeightMon", "WeightTue", "WeightWed", "WeightThu", "WeightFri", "WeightSat"][dow];
    return budInt(wrow[f]);
}
// 1店舗の日割り。M = 月予算（軽+普）、manual = { 日: 整数 }（手入力の日）。exact = 日ごとの値（小数）
function budPlan(month, M, wrow, manual) {
    var days = budMonthWeights(month).days, exact = {}, isManual = {}, O = 0, W = 0;
    days.forEach(function(d) { if (manual && manual[d] !== undefined && manual[d] !== null) { isManual[d] = true; O += manual[d]; } else W += budDayWeightBy(d, wrow); });
    var rest = M - O;
    days.forEach(function(d) { exact[d] = isManual[d] ? manual[d] : ((rest > 0 && W > 0) ? rest * budDayWeightBy(d, wrow) / W : 0); });
    return { days: days, exact: exact, manual: isManual, M: M, O: O, over: rest < 0 ? -rest : 0 };
}
// 表の予算データから、店舗ごとの日割りを作る { 店舗: plan }。月予算・日別予算のどちらかがある店舗だけ
function budStorePlans(bd) {
    var out = {};
    if (!bd || !Array.isArray(bd.rows)) return out;
    var M = {}, man = {}, rowsN = {};
    bd.rows.forEach(function(b) { var st = b.BudgetStore || "未所属"; M[st] = (M[st] || 0) + budInt(b.BudgetKei) + budInt(b.BudgetFu); rowsN[st] = (rowsN[st] || 0) + 1; });
    (bd.daily || []).forEach(function(b) {
        var st = b.BudgetStore || "未所属", d = String(b.TargetDate || "").slice(0, 10);
        if (d.slice(0, 7) !== bd.month || b.DailyBudget === null || b.DailyBudget === undefined || b.DailyBudget === "") return;
        (man[st] = man[st] || {})[d] = Math.round(budInt(b.DailyBudget)); rowsN[st] = (rowsN[st] || 0) + 1;
    });
    Object.keys(rowsN).forEach(function(st) { var p = budPlan(bd.month, M[st] || 0, budWeightRowFor(bd.weights, st, bd.month), man[st] || {}); p.rowCount = rowsN[st]; out[st] = p; });
    return out;
}
// 複数の店舗の日割りを足して、表示の整数にする。手入力の日はその整数、残りは累計を四捨五入して配る
function budScopeDisp(plans, days) {
    var disp = {}, autoCum = 0, prev = 0;
    days.forEach(function(d) {
        var man = 0, auto = 0;
        plans.forEach(function(p) { if (p.manual[d]) man += p.exact[d]; else auto += p.exact[d] || 0; });
        autoCum += auto; var r = Math.round(autoCum);
        disp[d] = man + (r - prev); prev = r;
    });
    return disp;
}
function budInt(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
function budStaffId(b) { var s = b && b.BudgetStaff; var id = s && typeof s === "object" ? s.id : s; return id ? String(id) : null; }
function budStaffName(b) { var id = budStaffId(b); if (!id) return null; return personMap[id] || b["BudgetStaff.Name"] || (b.BudgetStaff && b.BudgetStaff.name) || ("ID:" + id); }

// 日別予算の重み（シート 2025年9月の全店の日次予算の平均から）。日〜土、祝日は 125
var BUD_WD_WEIGHT = [186, 57, 50, 5, 49, 55, 168];
var BUD_HOLIDAY_WEIGHT = 125;
// 日本の祝日 2025〜2027（振替休日・国民の休日を含む）
var JP_HOLIDAYS = [
    "2025-01-01", "2025-01-13", "2025-02-11", "2025-02-23", "2025-02-24", "2025-03-20", "2025-04-29", "2025-05-03", "2025-05-04", "2025-05-05", "2025-05-06",
    "2025-07-21", "2025-08-11", "2025-09-15", "2025-09-23", "2025-10-13", "2025-11-03", "2025-11-23", "2025-11-24",
    "2026-01-01", "2026-01-12", "2026-02-11", "2026-02-23", "2026-03-20", "2026-04-29", "2026-05-03", "2026-05-04", "2026-05-05", "2026-05-06",
    "2026-07-20", "2026-08-11", "2026-09-21", "2026-09-22", "2026-09-23", "2026-10-12", "2026-11-03", "2026-11-23",
    "2027-01-01", "2027-01-11", "2027-02-11", "2027-02-23", "2027-03-21", "2027-03-22", "2027-04-29", "2027-05-03", "2027-05-04", "2027-05-05",
    "2027-07-19", "2027-08-11", "2027-09-20", "2027-09-23", "2027-10-11", "2027-11-03", "2027-11-23"
];
function budDayWeight(ds) {
    if (JP_HOLIDAYS.indexOf(ds) !== -1) return BUD_HOLIDAY_WEIGHT;
    var p = ds.split("-").map(Number);
    return BUD_WD_WEIGHT[new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay()];
}
// その月の日ごとの重みと合計。upto までの重みの合計も（今月の現時点予算用）
function budMonthWeights(month) {
    var days = [], w = {}, total = 0;
    var last = parseInt(lastDayOfMonth(month).slice(8), 10);
    for (var d = 1; d <= last; d++) { var ds = month + "-" + pad2(d); days.push(ds); w[ds] = budDayWeight(ds); total += w[ds]; }
    return { days: days, w: w, total: total };
}
// 現時点予算の割合: 過去の月は 1、今月は今日までの重み ÷ 月の重み、未来の月は 0
function budCurFrac(month, today) {
    var tm = today.slice(0, 7);
    if (month < tm) return 1;
    if (month > tm) return 0;
    var mw = budMonthWeights(month), s = 0;
    mw.days.forEach(function(d) { if (d <= today) s += mw.w[d]; });
    return mw.total ? s / mw.total : 0;
}

var STAT_BASES = ["j", "v_n", "sho", "ab", "jk", "rv", "rj", "tot_v", "n", "m", "c", "o2", "o3", "pk", "ct", "up", "tp", "ic", "rst", "ni", "nu", "hp", "fl", "aq", "tr", "ln", "l84", "r69", "r59", "r49", "r39", "r29", "low", "zn", "ar21", "ar22", "ar23", "ar24", "ar25", "ar_cnt", "del_cnt", "del_ar21", "del_ar22", "del_ar23", "del_ar24", "del_ar25", "ly_j", "ly_jt", "ly_vt"];
// budget_rows: 月次予算の行数（0 なら「未登録」）、daily_rows: その月の日次予算の行数
function createStats() { var s = { budget_n:0, budget_ar:0, budget_j:0, budget_j_k:0, budget_j_f:0, budget_t:0, budget_t_k:0, budget_t_f:0, budget_current:0, budget_rows:0, daily_rows:0, g_sls:0 }; STAT_BASES.forEach(function(b) { s[b + "_k"] = 0; s[b + "_f"] = 0; }); s.__recs = {}; return s; }

// --- 数字の内訳（どの案件を数えたか）を覚える。数は変えず、足した案件の参照を持つだけ ---
// s.__recs["j_k"] = Set(案件)。数の列（typeof number）だけを合算する処理には混ざらない
var STAT_KEYS = []; STAT_BASES.forEach(function(b) { STAT_KEYS.push(b + "_k", b + "_f"); });
// 件数が増えたら、同じ条件で += した金額の列にも同じ案件を覚えさせる（金額が 0 でも残す）
var REC_AMOUNT_WITH = { ar_cnt: ["ar21", "ar22", "ar23", "ar24", "ar25"], del_cnt: ["del_ar21", "del_ar22", "del_ar23", "del_ar24", "del_ar25"] };
function statsSnap(s) { var o = {}; for (var i = 0; i < STAT_KEYS.length; i++) o[STAT_KEYS[i]] = s[STAT_KEYS[i]]; return o; }
function noteRec(s, key, rec) { var m = s.__recs || (s.__recs = {}); var set = m[key] || (m[key] = new Set()); set.add(rec); }
function noteChanged(s, before, rec) {
    for (var i = 0; i < STAT_KEYS.length; i++) {
        var k = STAT_KEYS[i];
        if (s[k] === before[k]) continue;
        noteRec(s, k, rec);
        var suf = k.slice(-2), base = k.slice(0, -2);
        if (REC_AMOUNT_WITH[base]) REC_AMOUNT_WITH[base].forEach(function(a) { noteRec(s, a + suf, rec); });
    }
}
// 列の合計（店舗別・担当者別の「合計」）に、足し合わせた列の案件を集める
function mergeRecs(t, s) { var src = s.__recs || {}; Object.keys(src).forEach(function(k) { src[k].forEach(function(rec) { noteRec(t, k, rec); }); }); }

// コーティング・室内コート・防錆コート・ナビアップ・希望ナンバーは「空・✖・×・-None-」以外ならすべて獲得
function isGot(v) { if (v === null || v === undefined) return false; v = String(v).trim(); return ["", "✖", "×", "-None-"].indexOf(v) === -1; }
function catOf(rec) { return (rec.SyaryouCategory === "軽" || rec.SyaryouCategory === "軽自動車") ? "k" : "f"; }
function isCancelled(rec) { return rec.cancel === true || rec.cancel === "true"; }

function aggregate(s, rec) {
    var recBefore = statsSnap(s); // 内訳用（数の計算には使わない）
    var c = catOf(rec);
    var vD = (rec.VisitedDateTime || "").split('T')[0],
        cD = (rec.ClosingDay || ""),
        nD = (rec.nousyayoteibi || "").split('T')[0], // 納車日
        isCancel = isCancelled(rec);

    var st = curRange.start, ed = curRange.end;

    // 来店・成約集計
    if (vD && vD >= st && vD <= ed) {
        if (rec.FOrR === "初回" || rec.FOrR === "再来") s["tot_v_"+c]++; // 総接客数 = 新規 + 再来（初回/再来が空は入れない）
        if (rec.FOrR === "初回") s["v_n_"+c]++;
        if (rec.FOrR === "初回" && rec.Seated === "〇") s["sho_"+c]++; // 商談件数は初回の来店だけ
        if (rec.Option1 === "〇") s["ab_"+c]++;
        if (rec.FOrR === "初回" && cD === vD && !isCancel) s["jk_"+c]++;
        if (rec.FOrR === "再来") s["rv_"+c]++;
    }
    if (cD && cD >= st && cD <= ed && !isCancel) {
        s["j_"+c]++;
        if (rec.FOrR === "再来") s["rj_"+c]++;
        if (rec.HanbaiCategory === "新") s["n_"+c]++;
        if (rec.HanbaiCategory === "未") s["m_"+c]++;
        if (rec.HanbaiCategory === "中") s["c_"+c]++;
        if (rec.Option2 === "〇") s["o2_"+c]++;
        if (rec.Option3 === "〇") s["o3_"+c]++;
        if (["新車ﾊﾟｯｸ","未使用ﾊﾟｯｸ","中ｴｺﾊﾟｯｸ","中ｽﾀﾊﾟｯｸ","中ｱﾌﾟﾊﾟｯｸ"].indexOf(rec.Option15) !== -1) s["pk_"+c]++;
        if (isGot(rec.Option4)) s["ct_"+c]++;
        if (rec.Option5 === "〇") s["up_"+c]++;
        if (rec.Option16 === "〇") s["tp_"+c]++;
        if (isGot(rec.Option7)) s["ic_"+c]++;
        if (isGot(rec.Option8)) s["rst_"+c]++;
        if (rec.Option9 === "〇") s["ni_"+c]++;
        if (isGot(rec.Option10)) s["nu_"+c]++;
        if (isGot(rec.Option6)) s["hp_"+c]++;
        if (rec.BackCamera === "〇") s["fl_"+c]++;
        if (rec.Option17 === "〇") s["aq_"+c]++;
        if (["✖","✕","×","","null","-","無","なし"].indexOf(rec.TradeinCar || "") === -1) s["tr_"+c]++;
        if ((rec.PaymentCategory || "").indexOf("ローン") !== -1) { s["ln_"+c]++; var r = parseFloat(rec.arari17)||0, ct = parseInt(rec.arari16)||0; if (ct >= 84) s["l84_"+c]++; if (r >= 6 && r < 7) s["r69_"+c]++; else if (r >= 5 && r < 6) s["r59_"+c]++; else if (r >= 4 && r < 5) s["r49_"+c]++; else if (r >= 3 && r < 4) s["r39_"+c]++; else if (r >= 2.9 && r < 3) s["r29_"+c]++; else if (r > 0 && r < 2.9) s["low_"+c]++; }
        if (rec.Option14 === "〇") s["zn_"+c]++;
        s["ar21_"+c] += (parseFloat(rec.arari21)||0); s["ar22_"+c] += (parseFloat(rec.arari22)||0); s["ar23_"+c] += (parseFloat(rec.arari23)||0); s["ar24_"+c] += (parseFloat(rec.arari24)||0); s["ar25_"+c] += (parseFloat(rec.arari25)||0); s["ar_cnt_"+c]++;
    }

    // 納車集計 (納車日が期間内のもの)
    if (nD && nD >= st && nD <= ed && !isCancel) {
        s["del_cnt_" + c]++;
        s["del_ar21_" + c] += (parseFloat(rec.arari21) || 0); // 車両粗利
        s["del_ar23_" + c] += (parseFloat(rec.arari23) || 0); // ローンBK
        s["del_ar22_" + c] += (parseFloat(rec.arari22) || 0); // 下取粗利
        s["del_ar24_" + c] += (parseFloat(rec.arari24) || 0);
        s["del_ar25_" + c] += (parseFloat(rec.arari25) || 0);
    }
    noteChanged(s, recBefore, rec);
}

// 昨年: ly_j = 前年の同じ月の成約台数、ly_jt = 前年の開始日〜min(終了日,今日)の1年前までの成約台数、ly_vt = 同じ期間の新規接客数
function aggregateLY(s, rec) {
    var c = catOf(rec);
    var vD = (rec.VisitedDateTime || "").split('T')[0], cD = (rec.ClosingDay || ""), isCancel = isCancelled(rec);
    var R = curRange;
    var recBefore = statsSnap(s); // 内訳用（数の計算には使わない）
    if (cD && !isCancel) {
        if (cD >= R.lyMonthStart && cD <= R.lyMonthEnd) s["ly_j_"+c]++;
        if (cD >= R.lyStart && cD <= R.lyCutEnd) s["ly_jt_"+c]++;
    }
    if (vD && vD >= R.lyStart && vD <= R.lyCutEnd && rec.FOrR === "初回") s["ly_vt_"+c]++;
    noteChanged(s, recBefore, rec);
}

// 半角・全角の空白をすべて除いた名前（照合用）
function normName(v) { return String(v === null || v === undefined ? "" : v).replace(/[\s　]/g, ""); }
function personName(r) { if (r.ServicePerson && r.ServicePerson.id) return personMap[r.ServicePerson.id] || r["ServicePerson.Name"] || "ID:" + r.ServicePerson.id; return "未設定"; }

function renderAll(seq) {
    var loadingEl = document.getElementById('loading');
    loadFailed = false;
    curRange = getPeriod();
    var gS = {}, totalS = createStats(); staffStatsMaster = {}; storeStatsMaster = {}; staffStoreMap = {}; storeGroupMap = {}; var groupSet = new Set(), storeSet = new Set();

    // 担当者の列名は CRM 側の表記。予算側（"姓　名"）とは空白を除いた名前で照合する
    var staffKeyByNorm = {};
    function staffKey(name) { var n = normName(name); if (!n) return null; if (!staffKeyByNorm[n]) staffKeyByNorm[n] = name; return staffKeyByNorm[n]; }
    (allData || []).forEach(r => staffKey(personName(r)));

    // source: "cur"（今年の実績）/ "budget"（予算）/ "ly"（昨年）
    // 昨年だけに出てくる店舗・担当者の列は作らない（値は合計・グループ・既にある店舗の列には入れる）
    // 担当者の所属店は今年の実績で決め、予算・昨年では未設定のときだけ入れる
    function bucketsFor(st, pr, source) {
        var gr = groupOf(st);
        if(!gS[gr]) gS[gr] = createStats(); groupSet.add(gr);
        var list = [gS[gr], totalS];
        if (source !== "ly" && !storeStatsMaster[st]) storeStatsMaster[st] = createStats();
        if (storeStatsMaster[st]) { list.push(storeStatsMaster[st]); storeGroupMap[st] = gr; storeSet.add(st); }
        if (pr !== null) {
            if (source !== "ly" && !staffStatsMaster[pr]) staffStatsMaster[pr] = createStats();
            if (staffStatsMaster[pr]) { list.push(staffStatsMaster[pr]); if (source === "cur" || !staffStoreMap[pr]) staffStoreMap[pr] = st; }
        }
        return list;
    }

    (allData || []).forEach(r => { bucketsFor(r.ServiceStore || "未所属", staffKey(personName(r)), "cur").forEach(s => aggregate(s, r)); });

    // 予算・目標（販売予算）。担当者は id で照合し、名前は personMap（接客の担当者と同じ表記）から引いて staffKey に通す
    // 担当者が空の行は店舗・グループ・合計にだけ入る
    if (budgetDataGlobal && budgetDataGlobal.month === curRange.month && Array.isArray(budgetDataGlobal.rows)) {
        budgetDataGlobal.rows.forEach(function(b) {
            var st = b.BudgetStore || "未所属", nm = budStaffName(b);
            var bk = budInt(b.BudgetKei), bf = budInt(b.BudgetFu), tk = budInt(b.TargetKei), tf = budInt(b.TargetFu), bn = budInt(b.DeliveryBudget), bar = budInt(b.GrossProfitBudget);
            bucketsFor(st, nm ? staffKey(nm) : null, "budget").forEach(function(s) {
                s.budget_j += bk + bf; s.budget_j_k += bk; s.budget_j_f += bf;
                s.budget_t += tk + tf; s.budget_t_k += tk; s.budget_t_f += tf;
                s.budget_n += bn; s.budget_ar += bar; s.budget_rows++;
            });
        });
        // 現時点予算: 店舗ごとの日割り（手入力の日はその値、残りを店舗の曜日の重みで）の今日までの合計（担当者の列には入れない）
        var plans = budStorePlans(budgetDataGlobal);
        Object.keys(plans).forEach(function(st) {
            var p = plans[st], cur = 0;
            p.days.forEach(function(d) { if (d <= curRange.today) cur += p.exact[d]; });
            bucketsFor(st, null, "budget").forEach(function(s) { s.budget_current += cur; s.daily_rows += p.rowCount; });
        });
    }

    // 昨年は列を作ったあとに足す（昨年だけの店舗・担当者の列は作らない）
    if (lyAvailable) (lyData || []).forEach(r => { bucketsFor(r.ServiceStore || "未所属", staffKey(personName(r)), "ly").forEach(s => aggregateLY(s, r)); });

    // 500ms 後に読み込み表示を消す。その間に新しい読み込みが始まっていたら消さない
    setTimeout(() => { if (seq === undefined || !isStale(seq)) loadingEl.style.display = 'none'; }, 500);
    updateSelector('group-selector', orderKeys(Array.from(groupSet), "group"), '全グループ表示');
    updateSelector('store-selector', orderKeys(Array.from(storeSet), "store"), '全店舗表示');
    document.getElementById("group-table-container").innerHTML = buildTable(gS, "グループ名", totalS, "group");
    filterStoreByGroup(); filterStaffByStore();
    renderDaily();
    applyMyStoreOnce();
}

// 列の並び: グループは GROUP_ORDER、店舗はグループ順→グループ内の順（一覧に無い店舗は最後に名前順）、担当者は店舗の順→名前順
function byName(a, b) { return String(a).localeCompare(String(b), "ja"); }
function storeRank(st) { var i = STORE_ORDER.indexOf(st); return i === -1 ? STORE_ORDER.length : i; }
function orderKeys(keys, kind) {
    var arr = keys.slice();
    if (kind === "group") return arr.sort(function(a, b) { var ia = GROUP_ORDER.indexOf(a), ib = GROUP_ORDER.indexOf(b); if (ia === -1) ia = GROUP_ORDER.length; if (ib === -1) ib = GROUP_ORDER.length; return ia - ib || byName(a, b); });
    if (kind === "store") return arr.sort(function(a, b) { return storeRank(a) - storeRank(b) || byName(a, b); });
    if (kind === "staff") return arr.sort(function(a, b) { var sa = staffStoreMap[a], sb = staffStoreMap[b]; return storeRank(sa) - storeRank(sb) || byName(sa || "", sb || "") || byName(a, b); });
    return arr.sort();
}

function updateSelector(id, arr, def) { var sel = document.getElementById(id); var cur = sel.value; sel.innerHTML = '<option value="all">' + def + '</option>'; arr.forEach(v => { var o = document.createElement('option'); o.value = o.text = v; sel.add(o); }); if (cur && arr.indexOf(cur) !== -1) sel.value = cur; }

function buildTable(sum, title, totalS, kind) {
    var keys = orderKeys(Object.keys(sum), kind);
    // 店舗別は自分の店舗の列を一番左（合計の右）に
    if (kind === "store" && myStore && keys.indexOf(myStore) > 0) { keys.splice(keys.indexOf(myStore), 1); keys.unshift(myStore); }
    var h = "<table><thead><tr>";
    h += "<th class='sticky-col-item shop-header' style='position: sticky !important; z-index: 300 !important; top: 0; left: 0;'>" + (title || "KPI項目") + "</th>";
    h += "<th class='sticky-col-total shop-header' style='position: sticky !important; z-index: 290 !important; top: 0; left: 170px;'>合計</th>";
    for(var i=0; i<keys.length; i++) { h += "<th class='shop-header' style='position: sticky !important; z-index: 150 !important; top: 0;'>" + (kind === "store" && keys[i] === myStore ? "<span class='my-store-mark'>★</span>" : "") + keys[i] + "</th>"; }
    h += "</tr></thead><tbody>";

    // d2: 2段目（合計）の分母（キーを足し合わせる）。無ければ d と同じ
    const rowDef = [
        { sec: "予算・目標" },
        { lbl: "予算", m: "budget_j", type: "total_only", flag: "budget_rows", cls: "#ffffff" },
        { lbl: "目標", m: "budget_t", type: "total_only", flag: "budget_rows", cls: "#ffffff" },
        { lbl: "昨年実績", type: "ly_val", m: "ly_j", cls: "#ffffff" },
        { lbl: "現時点予算", m: "budget_current", type: "total_only", flag: "daily_rows", cls: "#ffffff" },
        { sec: "基本実績" },
        { lbl: "実績", m: "j", cls: "#ffe599" },
        { lbl: "達成率", type: "total_ratio", n: "j", d: "budget_current", flag: "daily_rows", cls: "#ffffff" },
        { lbl: "昨年実績(当日)", type: "ly_val", m: "ly_jt", cls: "#d9d2e9" }, { lbl: "昨年対比", type: "ly_ratio", n: "j", d: "ly_jt", cls: "#d9d2e9" },
        { lbl: "新規接客数", m: "v_n", cls: "#ffffff" }, { lbl: "昨年接客数(当日)", type: "ly_val", m: "ly_vt", cls: "#d9d2e9" }, { lbl: "商談件数", m: "sho", cls: "#ffffff" }, { lbl: "商談率", type: "ratio", n: "sho", d: "v_n", cls: "#ffe599" },
        { lbl: "AB数", m: "ab", cls: "#ffffff" }, { lbl: "AB率", type: "ratio", n: "ab", d: "sho", cls: "#ffffff" }, { lbl: "即決成約", m: "jk", cls: "#ffffff" }, { lbl: "即決率", type: "ratio", n: "jk", d: "v_n", cls: "#ffffff" },
        { lbl: "成約率", type: "ratio", n: "j", d: "v_n", cls: "#ffe599", redText: true }, { lbl: "昨年成約率", type: "ly_ratio", n: "ly_jt", d: "ly_vt", cls: "#d9d2e9" },
        { lbl: "再来接客数", m: "rv", cls: "#ffffff" }, { lbl: "再来店率", type: "custom_ratio", n: "rv", d_sub: ["v_n", "jk"], cls: "#ffffff" },
        { lbl: "再来成約", m: "rj", cls: "#ffffff" }, { lbl: "再来成約率", type: "custom_ratio", n: "rj", d_sub: ["v_n", "jk"], cls: "#ffffff" },
        { lbl: "総接客数", m: "tot_v", cls: "#ffffff" }, { lbl: "総成約率", type: "ratio", n: "j", d: "tot_v", cls: "#ffffff" },
        { lbl: "追客可能数", type: "diff_total", n: "v_n", d: "jk", cls: "#f4cccc" }, { lbl: "決着済", m: "empty", cls: "#f4cccc" }, { lbl: "決着率", m: "empty", cls: "#f4cccc" },
        { sec: "販売区分" }, { lbl: "新車", m: "n", cls: "#b6d7a8" }, { lbl: "獲得率", type: "ratio", n: "n", d: "j", cls: "#ffffff" }, { lbl: "未使用車", m: "m", cls: "#b6d7a8" }, { lbl: "獲得率", type: "ratio", n: "m", d: "j", cls: "#ffffff" }, { lbl: "中古車", m: "c", cls: "#b6d7a8" }, { lbl: "獲得率", type: "ratio", n: "c", d: "j", cls: "#ffffff" },
        { sec: "付帯品個別" },
        { lbl: "10年保証", m: "o2", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "o2", d: "j", d2: ["n", "m"], cls: "#ffffff" },
        { lbl: "中古保証", m: "o3", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "o3", d: "j", d2: ["c"], cls: "#ffffff" },
        { lbl: "パック", m: "pk", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "pk", d: "j", cls: "#ffffff" },
        { lbl: "コーティング", m: "ct", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "ct", d: "j", cls: "#ffffff" },
        { lbl: "UPグレード", m: "up", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "up", d: "j", cls: "#ffffff" },
        { lbl: "T-プレミアム", m: "tp", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "tp", d: "j", d2: ["n", "m"], cls: "#ffffff" },
        { lbl: "室内コート", m: "ic", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "ic", d: "j", cls: "#ffffff" },
        { lbl: "防錆コート", m: "rst", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "rst", d: "j", cls: "#ffffff" },
        { lbl: "ナビ取付", m: "ni", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "ni", d: "j", cls: "#ffffff" },
        { lbl: "ナビアップ", m: "nu", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "nu", d: "ni", cls: "#ffffff" },
        { lbl: "希望ナンバー", m: "hp", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "hp", d: "j", cls: "#ffffff" },
        { lbl: "フィルム", m: "fl", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "fl", d: "j", cls: "#ffffff" },
        { lbl: "アクアペル", m: "aq", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "aq", d: "j", cls: "#ffffff" },
        { lbl: "下取り", m: "tr", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "tr", d: "j", cls: "#ffffff" },
        { sec: "総付帯実績" }, { lbl: "車両総付帯", type: "sum", items: ["ct","up","tp","ic","rst","nu","fl","aq"], cls: "#a2c4c9" }, { lbl: "獲得率", type: "sum_ratio", items: ["ct","up","tp","ic","rst","nu","fl","aq"], d: "j", cls: "#ffffff" }, { lbl: "周辺総付帯", type: "sum", items: ["n","tr","ln","r69"], cls: "#a2c4c9" }, { lbl: "獲得率", type: "sum_ratio", items: ["n","tr","ln","r69"], d: "j", cls: "#ffffff" },
        { sec: "ローン実績" },
        { lbl: "ローン", m: "ln", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "ln", d: "j", cls: "#ffffff" },
        { lbl: "84回以上", m: "l84", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "l84", d: "j", d2: ["ln"], cls: "#ffffff" },
        { lbl: "6.9", m: "r69", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r69", d: "j", d2: ["ln"], cls: "#ffffff" },
        { lbl: "5.9", m: "r59", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r59", d: "j", d2: ["ln"], cls: "#ffffff" },
        { lbl: "4.9", m: "r49", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r49", d: "j", d2: ["ln"], cls: "#ffffff" },
        { lbl: "3.9", m: "r39", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r39", d: "j", d2: ["ln"], cls: "#ffffff" },
        { lbl: "2.9", m: "r29", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r29", d: "j", d2: ["ln"], cls: "#ffffff" },
        { lbl: "低金利", m: "low", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "low", d: "j", d2: ["ln"], cls: "#ffffff" },
        { lbl: "残価設定", m: "zn", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "zn", d: "j", d2: ["ln"], cls: "#ffffff" },
        { sec: "受注時(想定)" }, { lbl: "受注台数", type: "arari_val", val: "j", cls: "#ead1dc" }, { lbl: "＠車両粗利", type: "arari_avg", val: "ar21", cls: "#ead1dc" }, { lbl: "＠ローンBK", type: "arari_avg", val: "ar23", cls: "#ead1dc" }, { lbl: "＠下取粗利", type: "arari_avg", val: "ar22", cls: "#ead1dc" }, { lbl: "＠全部割(保証抜き)", type: "arari_avg", val: "ar24", cls: "#ead1dc" }, { lbl: "＠全部割(保証込み)", type: "arari_avg", val: "ar25", cls: "#ead1dc" }, { lbl: "総粗利(保証込み)", type: "arari_sum", val: "ar25", cls: "#ead1dc" },
        { sec: "納車着地粗利予測" },
        { lbl: "納車台数", type: "del_arari_val", val: "del_cnt", cls: "#d9ead3" },
        { lbl: "@車両粗利", type: "del_arari_avg", val: "del_ar21", cls: "#d9ead3" },
        { lbl: "@ローンBK", type: "del_arari_avg", val: "del_ar23", cls: "#d9ead3" },
        { lbl: "@下取粗利", type: "del_arari_avg", val: "del_ar22", cls: "#d9ead3" },
        { lbl: "@全部割(保証抜)", type: "del_arari_avg", val: "del_ar24", cls: "#d9ead3" },
        { lbl: "@全部割(保証込)", type: "del_arari_avg", val: "del_ar25", cls: "#d9ead3" },
        { lbl: "総粗利(保証抜)", type: "del_arari_sum", val: "del_ar24", cls: "#d9ead3" },
        { lbl: "総粗利(保証込)", type: "del_arari_sum", val: "del_ar25", cls: "#d9ead3" },
        { lbl: "納車台数予算", m: "budget_n", type: "total_only", flag: "budget_rows", cls: "#d9ead3" },
        { lbl: "達成率", type: "del_total_ratio", n: "del_cnt", d: "budget_n", flag: "budget_rows", cls: "#d9ead3" },
        { lbl: "総粗利保証込予算", m: "budget_ar", type: "total_only", flag: "budget_rows", cls: "#d9ead3" },
        { lbl: "達成率", type: "del_total_ratio", n: "del_ar25", d: "budget_ar", flag: "budget_rows", cls: "#d9ead3" }
    ];
    // セルを押したときに案件の一覧を出すため、この表の列と行の定義を覚えておく
    tableCtx[kind] = { sum: sum, keys: keys, totalS: totalS, rowDef: rowDef };

    for(var j=0; j<rowDef.length; j++){
        var r = rowDef[j];
        if(r.sec) {
            h += "<tr><td class='sticky-col-item section-row' style='position: sticky !important; z-index: 180 !important; left: 0; border-right: none;'>" + r.sec + "</td><td class='sticky-col-total section-row' style='position: sticky !important; z-index: 170 !important; left: 170px; border-left: none;'></td>";
            if (keys.length > 0) h += "<td colspan='" + keys.length + "' class='section-row' style='position: static; z-index: 1; border-left: none;'></td>";
            h += "</tr>";
        } else {
            h += "<tr><td class='sticky-col-item' style='background-color:"+(r.cls||"#fff")+"'>"+r.lbl+"</td>";
            h += markBudCell(markCell(renderCell(totalS, r, true), totalS, r, kind, j, -1), r, kind, null);
            for(var k=0; k<keys.length; k++) h += markBudCell(markCell(renderCell(sum[keys[k]], r, false), sum[keys[k]], r, kind, j, k), r, kind, keys[k]);
            h += "</tr>";
        }
    }
    h += "</tbody></table>"; return h;
}

// 率はシートに合わせて小数1桁（例 85.4%）。分母が 0 なら 0.0%（pctDash は「-」）
function fmtPct(n, d) { return (Math.round(n / d * 1000) / 10).toFixed(1) + "%"; }
function pct(n, d) { return d ? fmtPct(n, d) : "0.0%"; }
function pctDash(n, d) { return d ? fmtPct(n, d) : "-"; }
function sumKeys(s, keys, suf) { var t = 0; keys.forEach(function(k) { t += (s[k + suf] || 0); }); return t; }

function singleCell(c, cellStyle, val, extra) {
    return "<td class='"+c+"' style='"+cellStyle+"'><div class='cell-stack' style='align-items:center; font-weight:bold; font-size:12px;" + (extra || "") + "'>" + val + "</div></td>";
}
function twoTierCell(c, cellStyle, textClass, kVal, fVal, tVal) {
    return "<td class='"+c+"' style='"+cellStyle+"'><div class='cell-stack "+textClass+"'><div class='stack-upper' style='display:flex;'><div class='val-kei'>"+kVal+"</div><div class='val-fu'>"+fVal+"</div></div><div class='stack-lower' style='font-weight:bold;'>"+tVal+"</div></div></td>";
}

function renderCell(s, r, isT) {
    var kVal = "-", fVal = "-", tVal = "-", bg = r.cls || "#ffffff", textClass = r.redText ? "force-red" : "", c = isT ? "sticky-col-total " : "", cellStyle = "background-color:" + bg + " !important;";

    // 1. 予算・単一項目（予算の行が無ければ「未登録」）
    if(r.type === "total_only") {
        tVal = (r.flag && !s[r.flag]) ? "未登録" : Math.round(s[r.m] || 0).toLocaleString(); // 現時点予算は日割りで端数が出るので四捨五入して出す
        return singleCell(c, cellStyle, tVal, " color:#444;");
    }
    // 予算・目標（軽/普の二段。行が無ければ「未登録」）
    if(r.type === "budget_kf") {
        if (r.flag && !s[r.flag]) return singleCell(c, cellStyle, "未登録", " color:#444;");
        var bk = Math.round(s[r.m + "_k"] || 0), bf = Math.round(s[r.m + "_f"] || 0);
        return twoTierCell(c, cellStyle, "", bk.toLocaleString(), bf.toLocaleString(), (bk + bf).toLocaleString());
    }
    if(r.type === "total_dash") return singleCell(c, cellStyle, "-", " color:#444;");

    // 2. 達成率（分母の予算が未登録・0 なら「-」）
    if(r.type === "total_ratio" || r.type === "del_total_ratio") {
        var act = (s[r.n+"_k"] || 0) + (s[r.n+"_f"] || 0);
        var bud = Math.round(s[r.d] || 0); // 表示している予算（四捨五入した整数）で割る。0 なら「-」
        tVal = ((r.flag && !s[r.flag]) || bud <= 0) ? "-" : fmtPct(act, bud);
        return singleCell(c, cellStyle, tVal, "");
    }

    // 追客可能数 = 新規接客数 − 即決成約（合計1段）
    if(r.type === "diff_total") {
        tVal = ((s[r.n+"_k"] || 0) + (s[r.n+"_f"] || 0) - (s[r.d+"_k"] || 0) - (s[r.d+"_f"] || 0)).toLocaleString();
        return singleCell(c, cellStyle, tVal, "");
    }

    if(r.m === "empty") return twoTierCell(c, cellStyle, "", "-", "-", "-");

    // 昨年の行（前年の期間が CRM のデータより前なら「-」）
    if(r.type === "ly_val" || r.type === "ly_ratio") {
        if(!lyAvailable) return twoTierCell(c, cellStyle, "", "-", "-", "-");
        if(r.type === "ly_val") { var lk = s[r.m+"_k"] || 0, lf = s[r.m+"_f"] || 0; return twoTierCell(c, cellStyle, "", lk.toLocaleString(), lf.toLocaleString(), (lk + lf).toLocaleString()); }
        var nk = s[r.n+"_k"] || 0, nf = s[r.n+"_f"] || 0, dk = s[r.d+"_k"] || 0, df = s[r.d+"_f"] || 0;
        return twoTierCell(c, cellStyle, "", pctDash(nk, dk), pctDash(nf, df), pctDash(nk + nf, dk + df));
    }

    var isArari = r.type && r.type.startsWith("arari");
    var isDel = r.type && r.type.startsWith("del_arari");

    // --- 【受注・納車】二段構成セル（上段=合計、下段=軽/普） ---
    if(isArari || isDel) {
        var countK = isDel ? (s.del_cnt_k || 0) : (s.ar_cnt_k || 0);
        var countF = isDel ? (s.del_cnt_f || 0) : (s.ar_cnt_f || 0);

        if(r.type.endsWith("_val")) {
            kVal = (s[r.val+"_k"] || 0); fVal = (s[r.val+"_f"] || 0); tVal = (kVal + fVal).toLocaleString();
            kVal = kVal.toLocaleString(); fVal = fVal.toLocaleString();
        }
        else if(r.type.endsWith("_avg")) {
            var kA = countK ? Math.round((s[r.val+"_k"]||0)/countK) : 0;
            var fA = countF ? Math.round((s[r.val+"_f"]||0)/countF) : 0;
            var tA = (countK+countF) ? Math.round(((s[r.val+"_k"]||0)+(s[r.val+"_f"]||0))/(countK+countF)) : 0;
            kVal = kA.toLocaleString(); fVal = fA.toLocaleString(); tVal = tA.toLocaleString();
        }
        else if(r.type.endsWith("_sum")) {
            var kS = Math.round(s[r.val+"_k"] || 0), fS = Math.round(s[r.val+"_f"] || 0);
            tVal = (kS + fS).toLocaleString();
            kVal = kS.toLocaleString(); fVal = fS.toLocaleString();
        }

        var lowerBg = isDel ? "background:#ffffff !important;" : "background:transparent !important;";
        var upperBg = "background:transparent !important;";

        return "<td class='"+c+"' style='"+cellStyle+"'><div class='cell-stack'>" +
               "<div class='bg-sou-upper' style='"+upperBg+" border-bottom:1px dotted #999 !important; font-weight:bold;'>"+tVal+"</div>" +
               "<div class='bg-sou-lower' style='"+lowerBg+" color:#333; display:flex;'>" +
               "<div class='val-kei'>"+kVal+"</div><div class='val-fu'>"+fVal+"</div>" +
               "</div></div></td>";
    }

    // --- 【実績】三段構成セル ---
    if(r.lbl === "実績" && !r.type) {
        kVal = (s.j_k || 0).toLocaleString(); fVal = (s.j_f || 0).toLocaleString(); tVal = ((s.j_k || 0) + (s.j_f || 0)).toLocaleString();
        return "<td class='"+c+"' style='"+cellStyle+"'><div class='cell-stack'>" +
               "<div class='stack-label-3'><div style='width:50%;border-right:1px dotted #ccc;'>軽</div><div style='width:50%;'>普</div></div>" +
               "<div class='stack-values-3' style='background:#ffffff !important; display:flex;'><div class='val-kei'>"+kVal+"</div><div class='val-fu'>"+fVal+"</div></div>" +
               "<div class='stack-total-3' style='background:#ffffff !important; font-weight:bold;'>"+tVal+"</div></div></td>";
    }

    // --- その他 (通常の2段項目) ---
    if(r.type === "ratio") {
        var nk_k = s[r.n+"_k"] || 0, dk_k = s[r.d+"_k"] || 0, nk_f = s[r.n+"_f"] || 0, dk_f = s[r.d+"_f"] || 0;
        var d2 = r.d2 ? sumKeys(s, r.d2, "_k") + sumKeys(s, r.d2, "_f") : dk_k + dk_f;
        kVal = pct(nk_k, dk_k); fVal = pct(nk_f, dk_f); tVal = pct(nk_k + nk_f, d2);
    }
    else if(r.type === "custom_ratio") {
        var cn_k = s[r.n+"_k"], cd_k = (s[r.d_sub[0]+"_k"] - s[r.d_sub[1]+"_k"]), cn_f = s[r.n+"_f"], cd_f = (s[r.d_sub[0]+"_f"] - s[r.d_sub[1]+"_f"]);
        kVal = cd_k > 0 ? fmtPct(cn_k, cd_k) : "0.0%"; fVal = cd_f > 0 ? fmtPct(cn_f, cd_f) : "0.0%";
        tVal = (cd_k+cd_f) > 0 ? fmtPct(cn_k + cn_f, cd_k + cd_f) : "0.0%";
    }
    else if(r.type === "sum") {
        var sk = sumKeys(s, r.items, "_k"), sf = sumKeys(s, r.items, "_f");
        kVal = sk.toLocaleString(); fVal = sf.toLocaleString(); tVal = (sk + sf).toLocaleString();
    }
    else if(r.type === "sum_ratio") {
        var rk = sumKeys(s, r.items, "_k"), rf = sumKeys(s, r.items, "_f"), rdk = s[r.d+"_k"] || 0, rdf = s[r.d+"_f"] || 0;
        kVal = pct(rk, rdk); fVal = pct(rf, rdf); tVal = pct(rk + rf, rdk + rdf);
    }
    else { kVal = (s[r.m+"_k"] || 0).toLocaleString(); fVal = (s[r.m+"_f"] || 0).toLocaleString(); tVal = ((s[r.m+"_k"] || 0) + (s[r.m+"_f"] || 0)).toLocaleString(); }

    return twoTierCell(c, cellStyle, textClass, kVal, fVal, tVal);
}

function showPage(id) { document.querySelectorAll('.page-content').forEach(function(p){ p.classList.remove('active'); }); document.getElementById(id).classList.add('active'); document.querySelectorAll('.tab-btn').forEach(function(b){ b.classList.remove('active'); }); var btnId = 'btn-' + id.replace('-page', ''); var btn = document.getElementById(btnId); if(btn) btn.classList.add('active'); document.getElementById('group-filter-area').style.display = (id === 'store-page') ? 'block' : 'none'; document.getElementById('staff-filter-area').style.display = (id === 'staff-page') ? 'block' : 'none'; }
function filterStoreByGroup() { if (loadFailed) { document.getElementById("store-table-container").innerHTML = ""; return; } var g = document.getElementById('group-selector').value, f = {}, t = createStats(); Object.keys(storeStatsMaster).forEach(function(st){ if (g === "all" || storeGroupMap[st] === g) { f[st] = storeStatsMaster[st]; Object.keys(t).forEach(function(k){ if(typeof t[k] === 'number') t[k] += storeStatsMaster[st][k]; }); mergeRecs(t, storeStatsMaster[st]); } }); document.getElementById("store-table-container").innerHTML = buildTable(f, "店舗名", t, "store"); }
function filterStaffByStore() { if (loadFailed) { document.getElementById("staff-table-container").innerHTML = ""; return; } var s = document.getElementById('store-selector').value, f = {}, t = createStats(); Object.keys(staffStatsMaster).forEach(function(n){ if (s === "all" || staffStoreMap[n] === s) { f[n] = staffStatsMaster[n]; Object.keys(t).forEach(function(k){ if(typeof t[k] === 'number') t[k] += staffStatsMaster[n][k]; }); mergeRecs(t, staffStatsMaster[n]); } }); document.getElementById("staff-table-container").innerHTML = buildTable(f, "担当者名", t, "staff"); }

// ===== 日別の表（シート「進捗表(Gr込)」BX〜CM に合わせる） =====
// 数え方は aggregate / aggregateLY と同じ条件。表の数字（3つの表）には関わらない
var DAILY_WD = ["日", "月", "火", "水", "木", "金", "土"];
var dailyCtx = null;   // 押したときの一覧用: { days, scopeLabel, sets: { 列: [日ごとの案件の配列] }, lyOk: [日ごとに昨年を出すか] }
var DAILY_COLS = [
    { key: "date", lbl: "日" }, { key: "wd", lbl: "曜日" },
    { key: "vn", lbl: "新規来場", click: "vn" }, { key: "rv", lbl: "再来場", click: "rv" }, { key: "tot", lbl: "総来場", click: "tot" },
    { key: "bud", lbl: "予算" }, { key: "fc", lbl: "着地予想" }, { key: "j", lbl: "実績", click: "j" },
    { key: "bprog", lbl: "予算進捗" }, { key: "fprog", lbl: "着地予想進捗" },
    { key: "sho", lbl: "商談数", click: "sho" }, { key: "shoR", lbl: "商談率", click: "sho" }, { key: "jR", lbl: "成約率", click: "j" },
    { key: "lyvn", lbl: "昨年新規", click: "lyvn" }, { key: "lytot", lbl: "昨年総", click: "lytot" }, { key: "lyj", lbl: "昨年成約", click: "lyj" }
];
var DAILY_SET_LABEL = { vn: "新規来場", rv: "再来場", tot: "総来場", j: "実績", sho: "商談数", lyvn: "昨年新規", lytot: "昨年総", lyj: "昨年成約" };

function dailyDays(start, end) {
    var out = [];
    if (!start || !end || start > end) return out;
    var p = start.split("-").map(Number), d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
    for (var i = 0; i < 400; i++) { var s = d.toISOString().slice(0, 10); if (s > end) break; out.push(s); d.setUTCDate(d.getUTCDate() + 1); }
    return out;
}
function dailyWd(ds) { var p = ds.split("-").map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay(); }
function dailyMd(ds) { var p = ds.split("-").map(Number); return p[1] + "/" + p[2]; }
// 絞り込み: "all" / "g:グループ" / "s:店舗"。店舗は rec.ServiceStore（空は「未所属」）、グループは groupOf（既存の表と同じ）
function dailyInScope(st, f) {
    st = st || "未所属";
    if (!f || f === "all") return true;
    if (f.indexOf("g:") === 0) return groupOf(st) === f.slice(2);
    if (f.indexOf("s:") === 0) return st === f.slice(2);
    return true;
}
function dailyUpdateFilter() {
    var sel = document.getElementById("daily-filter");
    if (!sel) return "all";
    var cur = sel.value || "all";
    var stores = orderKeys(Object.keys(storeStatsMaster), "store");
    var groups = orderKeys(Array.from(new Set(stores.map(function(st) { return groupOf(st); }))), "group");
    var h = "<option value='all'>全店</option><optgroup label='グループ'>" + groups.map(function(g) { return "<option value='g:" + escHtml(g) + "'>" + escHtml(g) + "</option>"; }).join("") + "</optgroup>";
    h += "<optgroup label='店舗'>" + stores.map(function(st) { return "<option value='s:" + escHtml(st) + "'>" + escHtml(st) + "</option>"; }).join("") + "</optgroup>";
    sel.innerHTML = h;
    var ok = Array.prototype.some.call(sel.options, function(o) { return o.value === cur; });
    sel.value = ok ? cur : "all";
    return sel.value;
}

// 絞り込み範囲の日別予算。exact = 小数のまま、disp = 表示用の整数（累計を四捨五入した差で配るので、月の合計は月予算と一致する）
function dailyBudget(f) {
    var out = { ok: false, month: null, monthTotal: 0, rows: 0, exact: {}, disp: {} };
    if (!budgetDataGlobal || !Array.isArray(budgetDataGlobal.rows) || !budgetDataGlobal.month) return out;
    out.month = budgetDataGlobal.month;
    // 店舗ごとの日割り（手入力の日・店舗の曜日の重み）を、絞り込み範囲の店舗だけ足す
    var all = budStorePlans(budgetDataGlobal), plans = [];
    Object.keys(all).forEach(function(st) { if (!dailyInScope(st, f)) return; plans.push(all[st]); out.rows += all[st].rowCount; out.monthTotal += all[st].M; });
    out.ok = out.rows > 0;
    var days = budMonthWeights(out.month).days;
    days.forEach(function(d) { var x = 0; plans.forEach(function(p) { x += p.exact[d]; }); out.exact[d] = x; });
    out.disp = budScopeDisp(plans, days);
    return out;
}

function renderDaily() {
    var box = document.getElementById("daily-table-container");
    if (!box) return;
    if (loadFailed || !curRange) { box.innerHTML = ""; dailyCtx = null; return; }
    var f = dailyUpdateFilter();
    var days = dailyDays(curRange.start, curRange.end), idx = {};
    days.forEach(function(d, i) { idx[d] = i; });
    var sets = {}; ["vn", "rv", "tot", "j", "sho", "lyvn", "lytot", "lyj"].forEach(function(k) { sets[k] = days.map(function() { return []; }); });

    (allData || []).forEach(function(rec) {
        if (!dailyInScope(rec.ServiceStore, f)) return;
        var vD = (rec.VisitedDateTime || "").split("T")[0], cD = rec.ClosingDay || "";
        if (vD in idx) {
            var i = idx[vD];
            if (rec.FOrR === "初回") sets.vn[i].push(rec);
            if (rec.FOrR === "再来") sets.rv[i].push(rec);
            if (rec.FOrR === "初回" || rec.FOrR === "再来") sets.tot[i].push(rec);
            if (rec.FOrR === "初回" && rec.Seated === "〇") sets.sho[i].push(rec);
        }
        if ((cD in idx) && !isCancelled(rec)) sets.j[idx[cD]].push(rec);
    });

    // 昨年: 前年の同じ日（2/29 は 2/28）。今日より後・昨年が読めていない・CRM のデータより前の日は「-」
    var lyOk = days.map(function(d) { return lyAvailable && d <= curRange.today && shiftYear(d, -1) >= CRM_DATA_START; });
    var lyIdx = {};
    days.forEach(function(d, i) { if (!lyOk[i]) return; var ld = shiftYear(d, -1); (lyIdx[ld] = lyIdx[ld] || []).push(i); });
    if (Object.keys(lyIdx).length) (lyData || []).forEach(function(rec) {
        if (!dailyInScope(rec.ServiceStore, f)) return;
        var vD = (rec.VisitedDateTime || "").split("T")[0], cD = rec.ClosingDay || "";
        (lyIdx[vD] || []).forEach(function(i) {
            if (rec.FOrR === "初回") sets.lyvn[i].push(rec);
            if (rec.FOrR === "初回" || rec.FOrR === "再来") sets.lytot[i].push(rec);
        });
        if (!isCancelled(rec)) (lyIdx[cD] || []).forEach(function(i) { sets.lyj[i].push(rec); });
    });

    // 予算: 絞り込み範囲の店舗の月の成約予算（軽+普）を、曜日・祝日の重みで日に割る（dailyBudget）
    // 範囲に月予算の行が1つも無ければ「未登録」。読み込んだ月の外の日も「未登録」
    var db = dailyBudget(f), bud = days.map(function(d) { return db.disp[d] || 0; });
    var budOk = days.map(function(d) { return db.ok && d.slice(0, 7) === db.month; });

    var scopeLabel = f === "all" ? "全店" : f.slice(2);
    dailyCtx = { days: days, sets: sets, lyOk: lyOk, scopeLabel: scopeLabel };

    function n(k, i) { return sets[k][i].length; }
    function numCell(v, click, i, extraCls, style) {
        var attr = click ? " data-daily='" + click + "|" + i + "'" : "";
        return "<td class='daily-td" + (click ? " dl-cell" : "") + (extraCls ? " " + extraCls : "") + "'" + attr + (style ? " style='" + style + "'" : "") + "><div class='daily-val'>" + v + "</div></td>";
    }
    function progCell(v) { return v === null ? numCell("-", null, 0) : numCell(v.toLocaleString(), null, 0, v < 0 ? "daily-neg" : ""); }

    // 画面の幅に収める（日・曜日は細く、残りは均等。daily-fit）
    var h = "<table class='daily-table daily-fit'><colgroup><col class='dcol-day'><col class='dcol-wd'>" + DAILY_COLS.slice(2).map(function() { return "<col>"; }).join("") + "</colgroup><thead><tr>";
    DAILY_COLS.forEach(function(c, ci) { h += "<th class='shop-header" + (ci === 0 ? " daily-sticky" : "") + "'>" + c.lbl + "</th>"; });
    h += "</tr></thead><tbody>";
    var T = { vn: 0, rv: 0, tot: 0, j: 0, sho: 0, bud: 0, budAny: false, bprog: 0, lyvn: 0, lytot: 0, lyj: 0, lyAny: false };
    days.forEach(function(d, i) {
        var wd = dailyWd(d), wdCls = wd === 6 ? "daily-sat" : (wd === 0 ? "daily-sun" : "");
        var vn = n("vn", i), rv = n("rv", i), tot = n("tot", i), j = n("j", i), sho = n("sho", i);
        T.vn += vn; T.rv += rv; T.tot += tot; T.j += j; T.sho += sho;
        var bprog = null;
        if (budOk[i]) { T.bud += bud[i]; T.budAny = true; bprog = j - bud[i]; T.bprog += bprog; }
        h += "<tr>" + numCell(escHtml(dailyMd(d)), null, i, "daily-sticky " + wdCls) + numCell(DAILY_WD[wd], null, i, wdCls);
        h += numCell(vn.toLocaleString(), "vn", i) + numCell(rv.toLocaleString(), "rv", i) + numCell(tot.toLocaleString(), "tot", i);
        h += numCell(budOk[i] ? bud[i].toLocaleString() : "未登録", null, i) + numCell("-", null, i) + numCell(j.toLocaleString(), "j", i, "daily-j");
        h += progCell(bprog) + numCell("-", null, i);
        h += numCell(sho.toLocaleString(), "sho", i) + numCell(pctDash(sho, vn), "sho", i) + numCell(pctDash(j, vn), "j", i);
        if (lyOk[i]) {
            var a = n("lyvn", i), b = n("lytot", i), c = n("lyj", i);
            T.lyvn += a; T.lytot += b; T.lyj += c; T.lyAny = true;
            h += numCell(a.toLocaleString(), "lyvn", i, "daily-ly") + numCell(b.toLocaleString(), "lytot", i, "daily-ly") + numCell(c.toLocaleString(), "lyj", i, "daily-ly");
        } else h += numCell("-", null, i, "daily-ly") + numCell("-", null, i, "daily-ly") + numCell("-", null, i, "daily-ly");
        h += "</tr>";
    });
    // 合計（率は合計から計算）。押すと期間中の案件をまとめて出す（日 = -1）
    h += "<tr class='daily-total'>" + numCell("合計", null, -1, "daily-sticky") + numCell("", null, -1);
    h += numCell(T.vn.toLocaleString(), "vn", -1) + numCell(T.rv.toLocaleString(), "rv", -1) + numCell(T.tot.toLocaleString(), "tot", -1);
    h += numCell(T.budAny ? T.bud.toLocaleString() : "未登録", null, -1) + numCell("-", null, -1) + numCell(T.j.toLocaleString(), "j", -1, "daily-j");
    h += progCell(T.budAny ? T.bprog : null) + numCell("-", null, -1);
    h += numCell(T.sho.toLocaleString(), "sho", -1) + numCell(pctDash(T.sho, T.vn), "sho", -1) + numCell(pctDash(T.j, T.vn), "j", -1);
    if (T.lyAny) h += numCell(T.lyvn.toLocaleString(), "lyvn", -1, "daily-ly") + numCell(T.lytot.toLocaleString(), "lytot", -1, "daily-ly") + numCell(T.lyj.toLocaleString(), "lyj", -1, "daily-ly");
    else h += numCell("-", null, -1, "daily-ly") + numCell("-", null, -1, "daily-ly") + numCell("-", null, -1, "daily-ly");
    h += "</tr></tbody></table>";
    box.innerHTML = h;
}

// 日別のセルを押したときの一覧。率のセルは分子の案件
function dailyOpen(spec) {
    if (!dailyCtx) return;
    var p = String(spec || "").split("|"), key = p[0], i = parseInt(p[1], 10), ctx = dailyCtx;
    if (!ctx.sets[key] || isNaN(i) || i >= ctx.days.length) return;
    var list = [], seen = new Set();
    var idxs = i === -1 ? ctx.days.map(function(_, x) { return x; }) : [i];
    idxs.forEach(function(x) { ctx.sets[key][x].forEach(function(rec) { if (!seen.has(rec)) { seen.add(rec); list.push(rec); } }); });
    var byClose = key === "j" || key === "lyj";
    list.sort(function(a, b) {
        var da = byClose ? (a.ClosingDay || "") : (a.VisitedDateTime || "").split("T")[0], db = byClose ? (b.ClosingDay || "") : (b.VisitedDateTime || "").split("T")[0];
        if (da !== db) return da < db ? -1 : 1;
        return String(a.Name || "").localeCompare(String(b.Name || ""), "ja");
    });
    var dayLbl = i === -1 ? "合計（" + dailyMd(ctx.days[0]) + "〜" + dailyMd(ctx.days[ctx.days.length - 1]) + "）" : dailyMd(ctx.days[i]) + "(" + DAILY_WD[dailyWd(ctx.days[i])] + ")";
    var title = "日別（" + ctx.scopeLabel + "） ／ " + dayLbl + " ／ " + DAILY_SET_LABEL[key] + " ／ " + list.length.toLocaleString() + "件";
    var subs = [];
    if (key.indexOf("ly") === 0) subs.push("前年の同じ日付の案件です");
    ddShowList(title, subs, list, [], null);
}
document.addEventListener("click", function(e) {
    if (!e.target || !e.target.closest) return;
    var td = e.target.closest("td.dl-cell");
    if (td) dailyOpen(td.getAttribute("data-daily"));
});

// ===== 予算入力（販売予算 SalesBudgets を店舗ごとに入れる） =====
var BI_COLS = [
    { f: "BudgetKei", lbl: "成約予算" }, { f: "TargetKei", lbl: "成約目標" }, // 予算・目標は軽・普に分けない（合計を BudgetKei / TargetKei に入れる。BudgetFu / TargetFu は使わないが、入っていれば合計に足す）
    { f: "DeliveryBudget", lbl: "納車予算" }, { f: "GrossProfitBudget", lbl: "粗利予算（保証込）" }
];
var biState = { month: "", store: "", rows: [], seq: 0 };

function biPrevMonth(m) { var p = m.split("-").map(Number); var y = p[0], mo = p[1] - 1; if (mo === 0) { y--; mo = 12; } return y + "-" + pad2(mo); }
function biStoreList() {
    var base = (ddPicks && ddPicks.ServiceStore) ? ddPicks.ServiceStore.filter(function(v) { return v !== DD_NONE; }) : STORE_ORDER.slice();
    Object.keys(storeStatsMaster).forEach(function(st) { if (st !== "未所属" && base.indexOf(st) === -1) base.push(st); });
    return orderKeys(base, "store");
}
// 表の月（上の月の選択欄。無ければ今月）
function biTableMonth() {
    var main = document.getElementById("month-selector");
    return (main && main.value) || (curRange && curRange.month) || getToday().slice(0, 7);
}
// 月・店舗の select を作り、月は渡された月を選ぶ（select の選択肢を作ってから値を入れる。無い月は足す）
function biFillSelectors(month) {
    var ms = document.getElementById("bi-month"), ss = document.getElementById("bi-store");
    if (!ms || !ss) return;
    var main = document.getElementById("month-selector"), keep = ms.value;
    if (ms.options.length !== main.options.length) {
        ms.innerHTML = "";
        Array.prototype.forEach.call(main.options, function(o) { var x = document.createElement("option"); x.value = o.value; x.text = o.text; ms.add(x); });
    }
    var want = month || keep || biTableMonth();
    if (!Array.prototype.some.call(ms.options, function(o) { return o.value === want; })) { var x = document.createElement("option"); x.value = want; x.text = want.replace("-0", "/").replace("-", "/"); ms.add(x); }
    ms.value = want;
    var cur = ss.value;
    ss.innerHTML = "<option value=''>（店舗を選んでください）</option>" + biStoreList().map(function(st) { return "<option value='" + escHtml(st) + "'>" + escHtml(st) + "</option>"; }).join("");
    if (cur && Array.prototype.some.call(ss.options, function(o) { return o.value === cur; })) ss.value = cur;
}
// 表のセルから開く。月は表の月、店舗は渡されたもの（空なら選択を空で）
function openBudgetInput(store) {
    showPage("budget-page");
    biFillSelectors(biTableMonth());
    var ss = document.getElementById("bi-store");
    ss.value = "";
    if (store && Array.prototype.some.call(ss.options, function(o) { return o.value === store; })) ss.value = store;
    biLoad();
    if (ddPicks === null) ddLoadPicks().then(function() { biFillSelectors(); });
}
// 上のボタンから開く。月は表の月、店舗は前に選んだもの
function biOpenPage() {
    showPage("budget-page");
    var before = (document.getElementById("bi-month").value || "") + "|" + (document.getElementById("bi-store").value || "");
    biFillSelectors(biTableMonth());
    // 店舗がまだ選ばれていなければ、自分の店舗を最初の値に（人が選んだ後は変えない）
    var ssel = document.getElementById("bi-store");
    if (!ssel.value && myStore && !biMyStoreUsed && Array.prototype.some.call(ssel.options, function(o) { return o.value === myStore; })) { ssel.value = myStore; biMyStoreUsed = true; }
    var now = document.getElementById("bi-month").value + "|" + document.getElementById("bi-store").value;
    if (now !== before || !document.getElementById("budget-table-container").innerHTML) biLoad();
    if (ddPicks === null) ddLoadPicks().then(function() { biFillSelectors(); });
}
function biInfo(text) { var el = document.getElementById("bi-info"); if (el) el.textContent = text; }
function biStatus(text, kind) { var el = document.getElementById("bi-status"); if (el) { el.textContent = text; el.className = kind ? "bi-" + kind : ""; } }

// その店舗の担当者を接客から決める（Masters の所属店舗は空のため）。選んだ月の1日の90日前〜月末（今日より先は今日まで）に
// その店舗の接客に出てきた ServicePerson を、接客の件数の多い順に [{ id, name, count }]
async function biStaffFromServices(month, store) {
    var p = (month + "-01").split("-").map(Number), d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
    d.setUTCDate(d.getUTCDate() - 90);
    var from = d.toISOString().slice(0, 10), to = minStr(lastDayOfMonth(month), getToday());
    if (to < from) return [];
    var where = "ServiceStore = '" + coqlStr(store) + "' and VisitedDateTime between '" + from + "T00:00:00+09:00' and '" + to + "T23:59:59+09:00'";
    var recs = await fetchServices(where, null, "Services", "ServicePerson, ServicePerson.Name");
    var map = {}, list = [];
    recs.forEach(function(r) {
        var id = r.ServicePerson && typeof r.ServicePerson === "object" ? r.ServicePerson.id : r.ServicePerson;
        if (!id) return;
        id = String(id);
        if (!map[id]) { map[id] = { id: id, name: personMap[id] || r["ServicePerson.Name"] || (r.ServicePerson && r.ServicePerson.name) || ("ID:" + id), count: 0 }; list.push(map[id]); }
        map[id].count++;
    });
    list.sort(function(a, b) { return b.count - a.count || String(a.name).localeCompare(String(b.name), "ja"); });
    return list;
}
// Masters の status（id → 値）。50件ずつ。読めなければ空（＝誰も除かない）
async function biMastersStatus(ids) {
    var out = {};
    for (var i = 0; i < ids.length; i += 50) {
        var chunk = ids.slice(i, i + 50).filter(function(id) { return /^\d+$/.test(id); });
        if (!chunk.length) continue;
        try {
            var rows = await fetchServices("id in (" + chunk.join(", ") + ")", null, "Masters", "id, Name, status");
            rows.forEach(function(m) { if (m && m.id) out[String(m.id)] = String(m.status || ""); });
        } catch (e) { console.error("担当者の在籍の確認エラー:", e); }
    }
    return out;
}
async function biLoad() {
    var month = document.getElementById("bi-month").value, store = document.getElementById("bi-store").value;
    var box = document.getElementById("budget-table-container");
    var seq = ++biState.seq;
    biState = { month: month, store: store, rows: [], seq: seq };
    biAddClose(); // 保存していない「追加」の行は読み直すと消える
    if (!month || !store) { box.innerHTML = ""; biStatus(month ? "店舗を選ぶと表が出ます" : "", ""); return; }
    biInfo("");
    biStatus("担当者を読み込み中…", "");
    box.innerHTML = "";
    var budRows = null, staff = null, errs = [], planErr = null;
    var got = await Promise.all([
        querySalesBudgets(month, store).catch(function(e) { errs.push("予算の読み込みに失敗しました（" + errStr(e) + "）"); return null; }),
        biStaffFromServices(month, store).catch(function(e) { errs.push("担当者の読み込みに失敗しました（" + errStr(e) + "）"); return null; }),
        queryWeightRows(month, store).catch(function(e) { planErr = "曜日の重みを読めませんでした（" + errStr(e) + "）"; return null; }),
        queryDailyRows(month, store).catch(function(e) { planErr = "日別の予算を読めませんでした（" + errStr(e) + "）"; return null; })
    ]);
    if (seq !== biState.seq) return;
    budRows = got[0]; staff = got[1];
    biState.weights = got[2] || []; biState.daily = got[3] || []; biState.planErr = planErr;
    biState.wChain = Promise.resolve(); biState.dChain = Promise.resolve(); biState.busyKeys = {};
    if (budRows === null) { biStatus(errs.join(" ／ "), "err"); return; } // 予算が読めないまま入れると重複して作るので、表を出さない

    var rows = [], byStaff = {}, storeRow = null;
    function addRow(id, name, note) { var r = { staffId: id, name: name, note: note || "", rec: null, chain: Promise.resolve() }; rows.push(r); if (id) byStaff[id] = r; return r; }
    var budByStaff = {};
    budRows.forEach(function(b) { var id = budStaffId(b); if (id) budByStaff[id] = b; else storeRow = storeRow || b; });
    // 1. 直近90日にこの店舗の接客に出てきた担当者（件数の多い順）。Masters で退職の人は除く（その月に予算の行がある人は出す）
    var sIds = (staff || []).map(function(x) { return x.id; });
    Object.keys(budByStaff).forEach(function(id) { if (sIds.indexOf(id) === -1) sIds.push(id); });
    var status = sIds.length ? await biMastersStatus(sIds) : {};
    if (seq !== biState.seq) return;
    (staff || []).forEach(function(x) {
        var retired = status[x.id] === "退職";
        if (retired && !budByStaff[x.id]) return;
        addRow(x.id, x.name, retired ? "退職" : "");
    });
    var nRecent = rows.length;
    // 2. その月に予算の行があるのに上に出てこない人
    Object.keys(budByStaff).forEach(function(id) { if (!byStaff[id]) addRow(id, budStaffName(budByStaff[id]), status[id] === "退職" ? "退職" : ""); });
    var nBudOnly = rows.length - nRecent;
    // 画面写真でも分かるように、読み込んだ担当者の人数を出す
    biInfo("担当者 " + nRecent + "人（直近90日の接客から" + (staff === null ? "。読み込みに失敗" : "") + "）" + (nBudOnly ? "＋予算の行がある人 " + nBudOnly + "人" : ""));
    rows.forEach(function(r) { if (r.staffId && budByStaff[r.staffId]) r.rec = budByStaff[r.staffId]; });
    // 4. 店舗全体（担当者なし）
    var sr = addRow(null, "（担当者に割り振らない分）", "担当者の合計に上乗せして数えます"); sr.rec = storeRow;
    biState.rows = rows;
    biRender();
    biStatus(errs.length ? errs.join(" ／ ") : "", errs.length ? "err" : "");
}
function biValStr(v) { if (v === null || v === undefined || v === "") return ""; var n = parseFloat(v); return isNaN(n) ? String(v) : String(Math.round(n)); }
function biRender() {
    var box = document.getElementById("budget-table-container");
    var h = "<table class='daily-table bi-table'><thead><tr><th class='shop-header sticky-col-item'>担当者</th>";
    BI_COLS.forEach(function(c) { h += "<th class='shop-header'>" + escHtml(c.lbl) + "</th>"; });
    h += "</tr></thead><tbody>";
    biState.rows.forEach(function(r, ri) {
        h += "<tr><td class='sticky-col-item bi-name'>" + escHtml(r.name) + (r.note ? "<div class='bi-note'>" + escHtml(r.note) + "</div>" : "") + "</td>";
        BI_COLS.forEach(function(c, ci) {
            var v = biValStr(r.rec ? r.rec[c.f] : null);
            h += "<td class='bi-td'><input type='text' inputmode='numeric' class='bi-input' data-r='" + ri + "' data-c='" + ci + "' data-orig='" + escHtml(v) + "' value='" + escHtml(v) + "' aria-label='" + escHtml(r.name + " " + c.lbl) + "'><div class='dd-ed-msg'></div></td>";
        });
        h += "</tr>";
    });
    h += "<tr class='daily-total'><td class='sticky-col-item'>合計</td>";
    BI_COLS.forEach(function(c, ci) { h += "<td><div class='daily-val' data-total='" + ci + "'></div></td>"; });
    h += "</tr></tbody></table><div id='bi-plan'></div>";
    box.innerHTML = h;
    box.querySelectorAll("input.bi-input").forEach(function(el) {
        el.addEventListener("keydown", function(e) {
            if (e.key === "Enter") { e.preventDefault(); biSaveCell(el, true); }
            else if (e.key === "Escape" || e.key === "Esc") { e.preventDefault(); if (!el.__busy) { el.value = el.getAttribute("data-orig") || ""; ddEdMsg(el, "", ""); } }
        });
        el.addEventListener("blur", function() { biSaveCell(el, false); });
        el.addEventListener("paste", function(e) { biPaste(el, e); });
    });
    biTotals();
}
function biTotals() {
    BI_COLS.forEach(function(c, ci) {
        var t = 0, any = false;
        biState.rows.forEach(function(r) { var v = r.rec ? r.rec[c.f] : null; if (v !== null && v !== undefined && v !== "") { t += budInt(v); any = true; } });
        var el = document.querySelector("#budget-table-container [data-total='" + ci + "']");
        if (el) el.textContent = any ? Math.round(t).toLocaleString() : "";
    });
    biRenderPlan(); // 月予算が変わると日別の自動の値も変わる
}

// ===== 曜日の重み（この店舗）と 日別の予算（この店舗・この月） =====
var WEIGHT_DEFAULT = { WeightMon: 57, WeightTue: 50, WeightWed: 5, WeightThu: 49, WeightFri: 55, WeightSat: 168, WeightSun: 186, WeightHoliday: BUD_HOLIDAY_WEIGHT };
function biMonthM() { var M = 0; (biState.rows || []).forEach(function(r) { if (r.rec) M += budInt(r.rec.BudgetKei) + budInt(r.rec.BudgetFu); }); return M; }
function biManualMap() { var m = {}; (biState.daily || []).forEach(function(b) { var d = String(b.TargetDate || "").slice(0, 10); if (d && b.DailyBudget !== null && b.DailyBudget !== undefined && b.DailyBudget !== "") m[d] = Math.round(budInt(b.DailyBudget)); }); return m; }
function biDailyRow(d) { var x = null; (biState.daily || []).forEach(function(b) { if (String(b.TargetDate || "").slice(0, 10) === d) x = b; }); return x; }
function biRenderPlan() {
    var box = document.getElementById("bi-plan");
    if (!box || !biState.month || !biState.store) return;
    if (biState.planErr) { box.innerHTML = "<div class='bi-sec-title'>曜日の重み・日別の予算</div><div class='bi-warn'>" + escHtml(biState.planErr) + "。重複して作らないよう、ここでは入力できません。開き直してください</div>"; return; }
    // 入れかけの値とカーソルを保つ
    var keep = {}, act = document.activeElement, actKey = act && act.getAttribute ? act.getAttribute("data-k") : null;
    box.querySelectorAll("input[data-k]").forEach(function(el) { if (el.value !== (el.getAttribute("data-orig") || "")) keep[el.getAttribute("data-k")] = el.value; });
    var month = biState.month, wrow = budWeightRowFor(biState.weights, biState.store, month), busy = biState.busyKeys || {};
    var mLabel = month.replace("-", "/");
    var src = wrow ? String(wrow.BudgetMonth).slice(0, 7).replace("-", "/") + " から適用" : "全店共通の重み（未設定）";
    var h = "<div class='bi-sec-title'>曜日の重み（この店舗）<span class='bi-sec-sub'>" + escHtml(src) + "</span><button type='button' class='tab-btn' id='bi-w-reset'>全店共通に戻す</button></div>";
    h += "<table class='daily-table bi-table bi-w-table'><thead><tr>" + WEIGHT_LABELS.map(function(l) { return "<th class='shop-header'>" + l + "</th>"; }).join("") + "</tr></thead><tbody><tr>";
    WEIGHT_FIELDS_LIST.forEach(function(f) {
        var v = wrow ? (wrow[f] === null || wrow[f] === undefined || wrow[f] === "" ? "0" : biValStr(wrow[f])) : String(WEIGHT_DEFAULT[f]);
        h += "<td class='bi-td'><input type='text' inputmode='numeric' class='bi-input-x bi-w-input' data-k='w:" + f + "' data-f='" + f + "' data-orig='" + escHtml(v) + "' value='" + escHtml(v) + "'" + (busy["w:" + f] ? " disabled" : "") + " aria-label='重み " + escHtml(WEIGHT_LABELS[WEIGHT_FIELDS_LIST.indexOf(f)]) + "'></td>";
    });
    h += "</tr></tbody></table><div class='bi-note2'>定休日は 0 を入れてください。値を直して Enter で保存すると、" + escHtml(mLabel) + " から適用します（前の月の設定は残ります）</div>";

    var M = biMonthM(), man = biManualMap(), plan = budPlan(month, M, wrow, man), disp = budScopeDisp([plan], plan.days);
    h += "<div class='bi-sec-title'>日別の予算（この店舗・" + escHtml(mLabel) + "）<span class='bi-sec-sub'>月予算 " + M.toLocaleString() + "（成約予算の合計）を、手入力の日を除いて曜日の重みで割ります。手入力を空にすると自動に戻ります</span></div>";
    if (plan.over > 0) h += "<div class='bi-warn'>手入力の日の合計が月予算を超えています（差 " + plan.over.toLocaleString() + "）。手入力でない日は 0 になります</div>";
    h += "<table class='daily-table bi-table bi-d-table'><thead><tr><th class='shop-header'>日</th><th class='shop-header'>曜日</th><th class='shop-header'>重み</th><th class='shop-header'>自動</th><th class='shop-header'>手入力</th><th class='shop-header'>この日の予算</th></tr></thead><tbody>";
    var tAuto = 0, tMan = 0, tAll = 0;
    plan.days.forEach(function(d) {
        var wd = dailyWd(d), wdCls = wd === 6 ? "daily-sat" : (wd === 0 ? "daily-sun" : ""), isM = !!plan.manual[d];
        var mv = isM ? String(man[d]) : "";
        if (isM) tMan += man[d]; else tAuto += disp[d];
        tAll += disp[d];
        h += "<tr" + (isM ? " class='bi-manual'" : "") + "><td class='" + wdCls + "'><div class='daily-val'>" + dailyMd(d) + (JP_HOLIDAYS.indexOf(d) !== -1 ? "<span class='bi-hol'>祝</span>" : "") + "</div></td><td class='" + wdCls + "'><div class='daily-val'>" + DAILY_WD[wd] + "</div></td>";
        h += "<td><div class='daily-val'>" + budDayWeightBy(d, wrow) + "</div></td><td><div class='daily-val'>" + (isM ? "-" : disp[d].toLocaleString()) + "</div></td>";
        h += "<td class='bi-td'><input type='text' inputmode='numeric' class='bi-input-x bi-d-input' data-k='d:" + d + "' data-day='" + d + "' data-orig='" + escHtml(mv) + "' value='" + escHtml(mv) + "'" + (busy["d:" + d] ? " disabled" : "") + " aria-label='" + d + " の日別予算'></td>";
        h += "<td><div class='daily-val'>" + disp[d].toLocaleString() + "</div></td></tr>";
    });
    h += "<tr class='daily-total'><td><div class='daily-val'>合計</div></td><td></td><td></td><td><div class='daily-val'>" + tAuto.toLocaleString() + "</div></td><td><div class='daily-val'>" + tMan.toLocaleString() + "</div></td><td><div class='daily-val'>" + tAll.toLocaleString() + (tAll === M ? "" : "（月予算 " + M.toLocaleString() + "）") + "</div></td></tr>";
    h += "</tbody></table>";
    box.innerHTML = h;
    box.querySelectorAll("input[data-k]").forEach(function(el) {
        var k = el.getAttribute("data-k");
        if (k in keep) el.value = keep[k];
        var save = function(viaEnter) { return k.indexOf("w:") === 0 ? biSaveWeight(el, viaEnter) : biSaveDaily(el, viaEnter); };
        el.addEventListener("keydown", function(e) {
            if (e.key === "Enter") { e.preventDefault(); save(true); }
            else if (e.key === "Escape" || e.key === "Esc") { e.preventDefault(); el.value = el.getAttribute("data-orig") || ""; }
        });
        el.addEventListener("blur", function() { save(false); });
    });
    var rb = document.getElementById("bi-w-reset");
    if (rb) rb.addEventListener("click", biResetWeights);
    if (actKey) { var f = box.querySelector("input[data-k='" + actKey + "']"); if (f && !f.disabled) { try { f.focus(); } catch (e) {} } }
}
function biErrMsg(x) { return ddErrText(x); }
// 手元（予算入力と表の予算データ）の重み・日別の行を書き換え、表を作り直す
function biApplyPlanGlobal(kind, rec, removed) {
    var month = biState.month, bd = budgetDataGlobal;
    if (bd && Array.isArray(bd.rows) && (kind === "w" ? month <= bd.month : month === bd.month)) {
        var list = kind === "w" ? (bd.weights = bd.weights || []) : (bd.daily = bd.daily || []);
        var i = -1; list.forEach(function(x, j) { if (String(x.id) === String(rec.id)) i = j; });
        if (removed) { if (i !== -1) list.splice(i, 1); }
        else if (i === -1) list.push(Object.assign({}, rec)); else Object.assign(list[i], rec);
        var loadingEl = document.getElementById("loading");
        if (!(loadingEl && loadingEl.style.display !== "none")) { try { renderAll(); } catch (e) { console.error("表の再計算エラー:", e); } }
    }
}
// 重みを1つ保存する。この月の重みの行があれば更新、無ければ今効いている値（無ければ共通）を写して新しく作る
function biSaveWeight(el, viaEnter, overrides) {
    var f = el ? el.getAttribute("data-f") : null, k = el ? el.getAttribute("data-k") : "w:*";
    if (el && (!el.isConnected || el.disabled || (biState.busyKeys || {})[k])) return Promise.resolve(false);
    var vals = {};
    if (!overrides) {
        var v = biParse(el.value);
        if (v === undefined || (v !== null && v < 0)) { biStatus("重みは 0 以上の整数で入れてください。保存していません", "err"); return Promise.resolve(false); }
        if (v === null) v = 0; // 空は 0（定休日）
        if (String(v) === (el.getAttribute("data-orig") || "")) { el.value = el.getAttribute("data-orig") || ""; return Promise.resolve(false); }
        vals[f] = v;
    } else vals = overrides;
    var seq = biState.seq, month = biState.month, store = biState.store;
    biState.busyKeys[k] = true; if (el) el.disabled = true;
    biStatus("重みを保存しています…", "");
    var p = biState.wChain.then(async function() {
        var wrow = budWeightRowFor(biState.weights, store, month), res = null, err = null, rec = null;
        var mine = wrow && String(wrow.BudgetMonth).slice(0, 7) === month;
        try {
            if (mine) {
                var up = Object.assign({ id: wrow.id }, vals);
                res = await ZOHO.CRM.API.updateRecord({ Entity: "SalesBudgets", APIData: up, Trigger: [] });
            } else {
                var ins = { Name: "重み " + month + " " + store, BudgetType: BT_WEIGHT, BudgetStore: store, BudgetMonth: month + "-01" };
                WEIGHT_FIELDS_LIST.forEach(function(x) { ins[x] = wrow ? (wrow[x] === null || wrow[x] === undefined || wrow[x] === "" ? 0 : Math.round(budInt(wrow[x]))) : WEIGHT_DEFAULT[x]; });
                Object.assign(ins, vals);
                res = await ZOHO.CRM.API.insertRecord({ Entity: "SalesBudgets", APIData: ins, Trigger: [] });
                rec = ins;
            }
        } catch (e) { err = e; }
        var ok = !err && res && Array.isArray(res.data) && res.data[0] && res.data[0].code === "SUCCESS";
        if (ok && !mine) { var nid = res.data[0].details && res.data[0].details.id; if (!nid) ok = false; else { rec = Object.assign({ id: String(nid) }, rec); if (seq === biState.seq) biState.weights.push(rec); } }
        if (ok && mine) { Object.assign(wrow, vals); rec = wrow; }
        if (seq === biState.seq) {
            delete biState.busyKeys[k];
            if (ok) biStatus("曜日の重みを保存しました（" + month.replace("-", "/") + " から適用）", "ok");
            else { console.error("重みの保存エラー:", err || res); biStatus("重み: " + biErrMsg(err || res), "err"); }
            biRenderPlan();
            if (viaEnter && f) { var n = document.querySelector("#bi-plan input[data-k='" + k + "']"); if (n) n.focus(); }
        }
        if (ok) biApplyPlanGlobal("w", rec, false);
        return ok;
    });
    biState.wChain = p.catch(function() {});
    return p;
}
// 全店共通に戻す: 共通の値をこの月の重みの行に保存する（行は消さない）
function biResetWeights() {
    if (!biState.month || !biState.store) return;
    if (!window.confirm("この店舗の曜日の重みを、" + biState.month.replace("-", "/") + " から全店共通の値（月57 火50 水5 木49 金55 土168 日186 祝125）にします。よろしいですか？")) return;
    biSaveWeight(null, false, Object.assign({}, WEIGHT_DEFAULT));
}
// 日別予算を1日保存する。入れた → その日の行を作る／更新。空 → 行を消して自動に戻す（確認なし）
function biSaveDaily(el, viaEnter) {
    var d = el.getAttribute("data-day"), k = "d:" + d;
    if (!el.isConnected || el.disabled || (biState.busyKeys || {})[k]) return Promise.resolve(false);
    var v = biParse(el.value), orig = el.getAttribute("data-orig") || "";
    if (v === undefined || (v !== null && v < 0)) { biStatus("日別の予算は 0 以上の整数で入れてください。保存していません", "err"); return Promise.resolve(false); }
    var vs = v === null ? "" : String(v);
    if (vs === orig) { el.value = orig; return Promise.resolve(false); }
    var seq = biState.seq, month = biState.month, store = biState.store;
    biState.busyKeys[k] = true; el.disabled = true;
    biStatus("日別の予算を保存しています…", "");
    var p = biState.dChain.then(async function() {
        var cur = biDailyRow(d), res = null, err = null, rec = null, removed = false;
        try {
            if (v === null) {
                if (cur) { res = await ZOHO.CRM.API.deleteRecord({ Entity: "SalesBudgets", RecordID: cur.id }); removed = true; rec = cur; }
                else res = { data: [{ code: "SUCCESS" }] };
            } else if (cur) {
                res = await ZOHO.CRM.API.updateRecord({ Entity: "SalesBudgets", APIData: { id: cur.id, DailyBudget: v }, Trigger: [] });
                rec = cur;
            } else {
                rec = { Name: "日別 " + d + " " + store, BudgetType: BT_DAILY, BudgetStore: store, TargetDate: d, BudgetMonth: month + "-01", DailyBudget: v };
                res = await ZOHO.CRM.API.insertRecord({ Entity: "SalesBudgets", APIData: rec, Trigger: [] });
            }
        } catch (e) { err = e; }
        var ok = !err && res && Array.isArray(res.data) && res.data[0] && res.data[0].code === "SUCCESS";
        if (ok && v !== null && !cur) { var nid = res.data[0].details && res.data[0].details.id; if (!nid) ok = false; else { rec = Object.assign({ id: String(nid) }, rec); if (seq === biState.seq) biState.daily.push(rec); } }
        if (ok && v !== null && cur) cur.DailyBudget = v;
        if (ok && removed && seq === biState.seq) biState.daily = biState.daily.filter(function(x) { return x !== cur; });
        if (seq === biState.seq) {
            delete biState.busyKeys[k];
            if (ok) biStatus(dailyMd(d) + " の日別予算を" + (v === null ? "自動に戻しました" : "保存しました"), "ok");
            else { console.error("日別予算の保存エラー:", err || res); biStatus(dailyMd(d) + ": " + biErrMsg(err || res), "err"); }
            biRenderPlan();
            if (viaEnter) { var n = document.querySelector("#bi-plan input[data-k='" + k + "']"); if (n) n.focus(); }
        }
        if (ok && rec) biApplyPlanGlobal("d", rec, removed);
        return ok;
    });
    biState.dChain = p.catch(function() {});
    return p;
}
// 入力を数に。空は null、整数でなければ undefined
function biParse(s) {
    s = String(s === null || s === undefined ? "" : s).replace(/[,，\s　]/g, "").replace(/[０-９]/g, function(ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); }).replace(/^[－ー]/, "-").replace(/^¥|^￥/, "");
    if (s === "") return null;
    if (!/^-?\d+$/.test(s)) return undefined;
    return parseInt(s, 10);
}
function biName(r) { return biState.month + " " + biState.store + " " + (r.staffId ? r.name : "（店舗）"); }
// セルを1つ保存する。同じ行の保存は直列（最初の新規作成が終わってから次を送る）
function biSaveCell(el, viaEnter) {
    if (!el || el.__busy) return Promise.resolve(false);
    var r = biState.rows[parseInt(el.getAttribute("data-r"), 10)], col = BI_COLS[parseInt(el.getAttribute("data-c"), 10)];
    if (!r || !col) return Promise.resolve(false);
    var orig = el.getAttribute("data-orig") || "", v = biParse(el.value);
    if (v === undefined) { ddEdMsg(el, "整数で入力してください。保存していません", "err"); return Promise.resolve(false); }
    var vs = v === null ? "" : String(v);
    if (vs === orig) { if (el.value !== orig) el.value = orig; return Promise.resolve(false); }
    var seq = biState.seq, month = biState.month, store = biState.store;
    el.__busy = true; el.disabled = true;
    ddEdMsg(el, "保存しています…", "");
    var p = r.chain.then(async function() {
        var res = null, err = null, isNew = !(r.rec && r.rec.id);
        try {
            if (!isNew) {
                var up = { id: r.rec.id }; up[col.f] = v;
                res = await ZOHO.CRM.API.updateRecord({ Entity: "SalesBudgets", APIData: up, Trigger: [] });
            } else {
                var ins = { Name: biName(r), BudgetMonth: month + "-01", BudgetStore: store, BudgetType: BT_MONTH };
                if (r.staffId) ins.BudgetStaff = { id: r.staffId };
                ins[col.f] = v;
                res = await ZOHO.CRM.API.insertRecord({ Entity: "SalesBudgets", APIData: ins, Trigger: [] });
            }
        } catch (e) { err = e; }
        var ok = !err && res && Array.isArray(res.data) && res.data[0] && res.data[0].code === "SUCCESS";
        if (ok && isNew) {
            var nid = res.data[0].details && res.data[0].details.id;
            if (!nid) ok = false;
            else r.rec = { id: String(nid), Name: biName(r), BudgetMonth: month + "-01", BudgetStore: store, BudgetStaff: r.staffId ? { id: r.staffId, name: r.name } : null, "BudgetStaff.Name": r.staffId ? r.name : null };
        }
        if (ok) {
            r.rec[col.f] = v;
            el.setAttribute("data-orig", vs); el.value = vs;
            biApplyGlobal(month, r.rec);
            ddEdMsg(el, "保存しました", "ok");
        } else {
            console.error("予算の保存エラー:", err || res);
            el.value = orig;
            ddEdMsg(el, ddErrText(err || res), "err");
        }
        el.disabled = false; el.__busy = false;
        if (viaEnter && document.body.contains(el)) { try { el.focus(); } catch (e) {} }
        if (seq === biState.seq) biTotals();
        return ok;
    });
    r.chain = p.catch(function() {});
    return p;
}
// 保存できた予算を手元の予算データに入れ、表を作り直す（読み込み直しはしない）
function biApplyGlobal(month, rec) {
    if (!budgetDataGlobal || budgetDataGlobal.month !== month || !Array.isArray(budgetDataGlobal.rows)) return;
    var cur = null;
    budgetDataGlobal.rows.forEach(function(b) { if (String(b.id) === String(rec.id)) cur = b; });
    if (!cur) budgetDataGlobal.rows.push(rec); // 新規作成した行は同じものを持つ（以後の書き換えもそのまま効く）
    else if (cur !== rec) BI_COLS.forEach(function(c) { if (c.f in rec) cur[c.f] = rec[c.f]; });
    var loadingEl = document.getElementById("loading");
    if (loadingEl && loadingEl.style.display !== "none") return;
    try { renderAll(); } catch (e) { console.error("表の再計算エラー:", e); }
}
// Excel・スプレッドシートからの貼り付け（タブ・改行区切り）を、そのセルを左上として流し込む
function biPaste(el, e) {
    var text = (e.clipboardData || window.clipboardData) ? (e.clipboardData || window.clipboardData).getData("text") : "";
    if (!text) return;
    text = text.replace(/\r\n?/g, "\n").replace(/\n$/, "");
    if (text.indexOf("\t") === -1 && text.indexOf("\n") === -1) return; // 1セルだけは普通の入力と同じ
    e.preventDefault();
    var r0 = parseInt(el.getAttribute("data-r"), 10), c0 = parseInt(el.getAttribute("data-c"), 10), targets = [];
    text.split("\n").forEach(function(line, dr) {
        line.split("\t").forEach(function(cell, dc) {
            var t = document.querySelector("#budget-table-container input.bi-input[data-r='" + (r0 + dr) + "'][data-c='" + (c0 + dc) + "']");
            if (!t || t.__busy) return;
            t.value = cell.trim();
            if (t.value !== (t.getAttribute("data-orig") || "")) targets.push(t);
        });
    });
    biSaveMany(targets);
}
async function biSaveMany(targets) {
    var n = 0;
    for (var i = 0; i < targets.length; i++) { if (await biSaveCell(targets[i], false)) n++; }
    biStatus(targets.length ? n + "件のセルを保存しました" + (n < targets.length ? "（" + (targets.length - n) + "件は保存できませんでした）" : "") : "変わったセルはありません", n < targets.length ? "err" : "ok");
    return n;
}
// 前月の同じ店舗の行を、今月の空のセルにだけ入れる（上書きしない）
// ===== 担当者を追加（表にいない人の行を足す。保存するまでは手元だけ） =====
var biAddTimer = null, biAddSeq = 0;
function biAddOpen() {
    var pop = document.getElementById("bi-add-pop");
    if (!pop) return;
    if (!biState.month || !biState.store) { biStatus("先に月と店舗を選んでください", "err"); return; }
    pop.hidden = false;
    var q = document.getElementById("bi-add-q");
    q.value = ""; document.getElementById("bi-add-list").innerHTML = "<div class='bi-add-msg'>名前の一部を2文字以上入れてください</div>";
    q.focus();
}
function biAddClose() { var pop = document.getElementById("bi-add-pop"); if (pop && !pop.hidden) { pop.hidden = true; biAddSeq++; clearTimeout(biAddTimer); } }
function biAddInput() {
    clearTimeout(biAddTimer);
    var q = document.getElementById("bi-add-q").value.trim(), list = document.getElementById("bi-add-list");
    if (q.length < 2) { biAddSeq++; list.innerHTML = "<div class='bi-add-msg'>名前の一部を2文字以上入れてください</div>"; return; }
    list.innerHTML = "<div class='bi-add-msg'>探しています…</div>";
    biAddTimer = setTimeout(function() { biAddSearch(q); }, 300); // 入力が止まって 0.3 秒後に読む
}
async function biAddSearch(q) {
    var seq = ++biAddSeq, list = document.getElementById("bi-add-list"), rows = null, err = null;
    var like = coqlStr(q).replace(/%/g, "").replace(/_/g, "");
    try { rows = (await coqlPage("Category = '担当者' and Name like '%" + like + "%'", 0, "Masters", "id, Name, status")).rows; }
    catch (e) { err = e; }
    if (seq !== biAddSeq) return;
    if (err) { list.innerHTML = "<div class='bi-add-msg bi-err'>読み込めませんでした（" + escHtmlerrStr(err) + "）</div>"; return; }
    var shown = {}; biState.rows.forEach(function(r) { if (r.staffId) shown[r.staffId] = true; });
    var cand = (rows || []).filter(function(m) { return m && m.id && String(m.status || "") !== "退職"; });
    cand.sort(function(a, b) { return String(a.Name || "").localeCompare(String(b.Name || ""), "ja"); });
    if (!cand.length) { list.innerHTML = "<div class='bi-add-msg'>見つかりません</div>"; return; }
    list.innerHTML = cand.map(function(m) {
        var id = String(m.id), on = shown[id];
        return "<div class='bi-add-item" + (on ? " bi-add-on" : "") + "' data-id='" + escHtml(id) + "' data-name='" + escHtml(m.Name || "") + "'>" + escHtml(m.Name || ("ID:" + id)) + (on ? "<span class='bi-add-tag'>表示中</span>" : "") + "</div>";
    }).join("");
}
// 選んだ人の行を「（担当者に割り振らない分）」の直前に足す（入力欄は空、注記「追加」）
function biAddPick(id, name) {
    if (!id || biState.rows.some(function(r) { return r.staffId === id; })) return;
    var r = { staffId: id, name: name || ("ID:" + id), note: "追加", rec: null, chain: Promise.resolve() };
    var at = -1; biState.rows.forEach(function(x, i) { if (at === -1 && !x.staffId) at = i; });
    if (at === -1) biState.rows.push(r); else biState.rows.splice(at, 0, r);
    biAddClose();
    biRender();
    var el = document.querySelector("#budget-table-container input.bi-input[data-r='" + biState.rows.indexOf(r) + "'][data-c='0']");
    if (el) el.focus();
    biStatus("「" + r.name + "」を足しました。値を入れて Enter で保存してください", "ok");
}
document.addEventListener("click", function(e) {
    var pop = document.getElementById("bi-add-pop");
    if (!pop || pop.hidden || !e.target || !e.target.closest) return;
    var item = e.target.closest(".bi-add-item");
    if (item && pop.contains(item)) { if (!item.classList.contains("bi-add-on")) biAddPick(item.getAttribute("data-id"), item.getAttribute("data-name")); return; }
    if (!pop.contains(e.target) && !e.target.closest("#bi-add")) biAddClose(); // 外側を押したら閉じる
});
document.addEventListener("keydown", function(e) { if (e.key === "Escape" || e.key === "Esc") biAddClose(); });

async function biCopyPrev() {
    if (!biState.month || !biState.store) return;
    var seq = biState.seq, prev = biPrevMonth(biState.month), prevRows;
    biStatus("前月（" + prev + "）を読み込み中…", "");
    try { prevRows = await querySalesBudgets(prev, biState.store); } catch (e) { biStatus("前月の予算を読めませんでした（" + errStr(e) + "）", "err"); return; }
    if (seq !== biState.seq) return;
    var plan = [], skipped = 0;
    prevRows.forEach(function(b) {
        var id = budStaffId(b), ri = -1;
        biState.rows.forEach(function(r, i) { if (ri === -1 && (id ? r.staffId === id : !r.staffId)) ri = i; });
        if (ri === -1) { skipped++; return; }
        BI_COLS.forEach(function(c, ci) {
            var pv = b[c.f];
            if (pv === null || pv === undefined || pv === "") return;
            var t = document.querySelector("#budget-table-container input.bi-input[data-r='" + ri + "'][data-c='" + ci + "']");
            if (!t || (t.getAttribute("data-orig") || "") !== "" || t.value.trim() !== "") return; // 入っているセルは上書きしない
            plan.push({ el: t, v: biValStr(pv) });
        });
    });
    if (!plan.length) { biStatus("前月から入れるセルはありません（空のセルが無いか、前月の予算がありません）" + (skipped ? "。前月の " + skipped + " 人はこの表にいないので入れていません" : ""), ""); return; }
    if (!window.confirm("前月（" + prev + "）の予算を、空のセル " + plan.length + " 件に入れて保存します（入っているセルは変えません）" + (skipped ? "。前月の " + skipped + " 人はこの表にいないので入れません" : "") + "。よろしいですか？")) { biStatus("", ""); return; }
    plan.forEach(function(x) { x.el.value = x.v; });
    await biSaveMany(plan.map(function(x) { return x.el; }));
}
document.addEventListener("click", function(e) {
    if (!e.target || !e.target.closest) return;
    var td = e.target.closest("td.bud-cell");
    if (td) openBudgetInput(td.getAttribute("data-bud") || "");
});

// ===== 数字のセルを押すと、その数字に数えた案件（接客）の一覧を出す =====
// 表の数字の計算には関わらない。aggregate / aggregateLY が覚えた s.__recs を読むだけ
var tableCtx = {};      // kind（group / store / staff）→ その表の列・行の定義
var dealNameMap = {};   // 商談 id → 商談名（お客様名）。COQL で名前が返らなかった分だけ、一覧を開いたときに読む
var ddSeq = 0;          // 一覧を開いた番号。閉じた・開き直した後に、前の読み込みの結果を書かない
var DD_KIND_LABEL = { group: "グループ", store: "店舗", staff: "担当者" };
var DD_VISIT_KEYS = ["tot_v", "v_n", "sho", "ab", "jk", "rv", "ly_vt"]; // 来店日で数えた列（ほかは成約日、del_ は納車日）
var DD_AMOUNT_FIELD = { ar21: "arari21", ar22: "arari22", ar23: "arari23", ar24: "arari24", ar25: "arari25", del_ar21: "arari21", del_ar22: "arari22", del_ar23: "arari23", del_ar24: "arari24", del_ar25: "arari25" };
var DD_AMOUNT_LABEL = { arari21: "車両粗利", arari22: "下取粗利", arari23: "ローンBK", arari24: "全部割(保証抜き)", arari25: "全部割(保証込み)" };

function escHtml(v) { return String(v === null || v === undefined ? "" : v).replace(/[&<>"']/g, function(ch) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]; }); }
function recSet(s, key) { return (s && s.__recs && s.__recs[key]) || null; }
function cntKF(s, key, cat) { var k = s[key + "_k"] || 0, f = s[key + "_f"] || 0; return cat === "k" ? k : (cat === "f" ? f : k + f); }

// 行の定義から「どの列の案件を出すか」を決める。案件に結びつかないセル（予算・目標・「-」・未登録・追客可能数）は null
function ddSpec(s, r) {
    if (!r || r.sec || r.m === "empty") return null;
    var t = r.type;
    if (t === "total_only" || t === "total_dash" || t === "diff_total") return null;
    if ((t === "ly_val" || t === "ly_ratio") && !lyAvailable) return null;
    if (t === "total_ratio" || t === "del_total_ratio") {
        if ((r.flag && !s[r.flag]) || Math.round(s[r.d] || 0) <= 0) return null; // 表示が「-」（renderCell と同じく四捨五入した予算で判定）
        return { nums: [r.n], amount: DD_AMOUNT_FIELD[r.n] || null, denom: function() { return { label: "分母（予算）", value: Math.round(s[r.d] || 0).toLocaleString(), unit: "" }; } };
    }
    if (t && (t.indexOf("arari") === 0 || t.indexOf("del_arari") === 0)) {
        var isDel = t.indexOf("del_") === 0;
        var cnt = /_val$/.test(t) ? r.val : (isDel ? "del_cnt" : "ar_cnt");
        return { nums: [cnt], amount: DD_AMOUNT_FIELD[r.val] || null };
    }
    if (t === "ratio") return { nums: [r.n], denom: function(cat) { return { label: "分母", value: (cat === "t" && r.d2) ? sumKeys(s, r.d2, "_k") + sumKeys(s, r.d2, "_f") : cntKF(s, r.d, cat), unit: "件" }; } };
    if (t === "custom_ratio") return { nums: [r.n], denom: function(cat) { return { label: "分母", value: cntKF(s, r.d_sub[0], cat) - cntKF(s, r.d_sub[1], cat), unit: "件" }; } };
    if (t === "ly_ratio") return { nums: [r.n], denom: function(cat) { return { label: "分母", value: cntKF(s, r.d, cat), unit: "件" }; } };
    if (t === "sum") return { nums: r.items, items: true };
    if (t === "sum_ratio") return { nums: r.items, items: true, denom: function(cat) { return { label: "分母", value: cntKF(s, r.d, cat), unit: "件" }; } };
    if (t === "ly_val" || !t) return r.m ? { nums: [r.m] } : null;
    return null;
}
function ddHasRecs(s, sp) { for (var i = 0; i < sp.nums.length; i++) { var a = recSet(s, sp.nums[i] + "_k"), b = recSet(s, sp.nums[i] + "_f"); if ((a && a.size) || (b && b.size)) return true; } return false; }

// 押せるセルに印（data-dd="kind|行|列"、合計列は -1）と class を付ける。セルの中身・大きさは変えない
function markCell(html, s, r, kind, j, k) {
    var sp = ddSpec(s, r);
    if (!sp || !ddHasRecs(s, sp)) return html;
    return html.replace("<td class='", "<td data-dd='" + kind + "|" + j + "|" + k + "' class='dd-cell ");
}

// 予算・目標・現時点予算・納車台数予算・総粗利保証込予算のセルは、押すと予算入力をその月・その店舗で開く
// 店舗の列はその店舗、担当者の列はその人の店舗、グループ・合計の列は店舗を空で
var BUDGET_ROW_MS = ["budget_j", "budget_t", "budget_current", "budget_n", "budget_ar"];
function markBudCell(html, r, kind, key) {
    if (!r || BUDGET_ROW_MS.indexOf(r.m) === -1) return html;
    var st = "";
    if (key !== null && key !== undefined) { if (kind === "store") st = key; else if (kind === "staff") st = staffStoreMap[key] || ""; }
    return html.replace("<td class='", "<td data-bud='" + escHtml(st) + "' class='bud-cell ");
}

// 案件 → その案件を数えた列（sum の行では複数になる）
function ddCollect(s, sp, cat) {
    var sufs = cat === "t" ? ["_k", "_f"] : ["_" + cat], map = new Map();
    sp.nums.forEach(function(key) { sufs.forEach(function(suf) { var set = recSet(s, key + suf); if (set) set.forEach(function(rec) { var a = map.get(rec); if (!a) { a = []; map.set(rec, a); } if (a.indexOf(key) === -1) a.push(key); }); }); });
    return map;
}
function ddKeyLabel(rowDef, key) {
    var i, x;
    for (i = 0; i < rowDef.length; i++) { x = rowDef[i]; if (x.m === key && (!x.type || x.type === "ly_val")) return x.lbl; }
    for (i = 0; i < rowDef.length; i++) { x = rowDef[i]; if (x.val === key && /_(val|sum)$/.test(x.type || "")) return x.lbl; }
    return key;
}
function ddItemLabel(rowDef, key) { var l = ddKeyLabel(rowDef, key); return /^[\d.]+$/.test(l) ? "金利" + l : l; }
function ddRowLabel(rowDef, r) {
    var ref = null;
    if (["ratio", "custom_ratio", "ly_ratio", "total_ratio", "del_total_ratio"].indexOf(r.type) !== -1) ref = ddKeyLabel(rowDef, r.n);
    if (r.type === "sum_ratio") rowDef.forEach(function(x) { if (!ref && x.type === "sum" && x.items && x.items.join() === r.items.join()) ref = x.lbl; });
    return (ref && ref !== r.lbl) ? r.lbl + "（" + ref + "）" : r.lbl;
}
function ddDateOf(rec, key) {
    if (DD_VISIT_KEYS.indexOf(key) !== -1) return (rec.VisitedDateTime || "").split("T")[0];
    if (key.indexOf("del_") === 0) return (rec.nousyayoteibi || "").split("T")[0];
    return rec.ClosingDay || "";
}
function ddDealName(rec) { var d = rec.Deal_Name; if (!d) return ""; if (typeof d === "object") return d.name || dealNameMap[d.id] || ""; return String(d); }
function ddNum(v) { return parseFloat(v) || 0; }

function ddOpenFromCell(dd, cat) {
    var p = String(dd || "").split("|"), ctx = tableCtx[p[0]];
    if (!ctx) return;
    var r = ctx.rowDef[parseInt(p[1], 10)], k = parseInt(p[2], 10);
    var s = k === -1 ? ctx.totalS : ctx.sum[ctx.keys[k]];
    if (!r || !s) return;
    var sp = ddSpec(s, r);
    if (!sp) return;
    if (cat !== "k" && cat !== "f") cat = "t";

    var map = ddCollect(s, sp, cat), sortKey = sp.nums[0];
    var list = Array.from(map.keys());
    list.sort(function(a, b) {
        var da = ddDateOf(a, sortKey), db = ddDateOf(b, sortKey);
        if (da !== db) { if (!da) return 1; if (!db) return -1; return da < db ? -1 : 1; }
        return String(a.Name || "").localeCompare(String(b.Name || ""), "ja");
    });

    // 見出し: どの列・どの行・件数
    var colName = k === -1 ? "合計" : ctx.keys[k];
    var colKind = k === -1 ? "" : (DD_KIND_LABEL[p[0]] || "") + ": ";
    var catLabel = cat === "k" ? "軽" : (cat === "f" ? "普" : "軽・普の合計");
    var title = colKind + colName + " ／ " + ddRowLabel(ctx.rowDef, r) + " ／ " + catLabel + " ／ " + list.length.toLocaleString() + "件";

    var itemTotal = 0; map.forEach(function(a) { itemTotal += a.length; });
    var subs = [];
    if (sp.denom) {
        var dn = sp.denom(cat);
        var numTxt = sp.items ? "分子 " + itemTotal.toLocaleString() + "個（案件 " + list.length.toLocaleString() + "件）" : "分子 " + list.length.toLocaleString() + "件";
        subs.push(numTxt + " / " + dn.label + " " + (typeof dn.value === "number" ? dn.value.toLocaleString() : dn.value) + dn.unit);
    } else if (sp.items) {
        subs.push("付帯の数 " + itemTotal.toLocaleString() + "個（案件 " + list.length.toLocaleString() + "件）");
    }
    if (sp.amount) {
        var amt = 0; list.forEach(function(rec) { amt += ddNum(rec[sp.amount]); });
        subs.push(DD_AMOUNT_LABEL[sp.amount] + "の合計 " + Math.round(amt).toLocaleString() + " ／ 1台あたり " + (list.length ? Math.round(amt / list.length) : 0).toLocaleString());
    }
    var extraHead = [];
    if (sp.items) extraHead.push("数えた項目");
    if (sp.amount) extraHead.push(DD_AMOUNT_LABEL[sp.amount]);
    ddShowList(title, subs, list, extraHead, function(rec) {
        var x = "";
        if (sp.items) x += "<td>" + escHtml(map.get(rec).map(function(key) { return ddItemLabel(ctx.rowDef, key); }).join("・")) + "</td>";
        if (sp.amount) x += "<td class='dd-num'>" + escHtml(ddNum(rec[sp.amount]).toLocaleString()) + "</td>";
        return x;
    });
}

// 一覧（モーダル）を出す。extraHead / extraCell は右端に足す列（無ければ空）
function ddShowList(title, subs, list, extraHead, extraCell) {
    subs = subs.slice();
    subs.push("行を押すと接客のレコードを新しいタブで開きます");
    subs.push("管理番号・日付・店舗・車両区分はここで直せます（Enter で保存）");

    ddCurRecs = {};
    list.forEach(function(rec) { ddCurRecs[String(rec.id)] = rec; });
    var h = "<table class='dd-table'><thead><tr><th>接客番号</th><th>管理番号</th><th>お客様名</th><th>来店日</th><th>成約日</th><th>納車日</th><th>店舗</th><th>担当者</th><th>車両区分</th>";
    (extraHead || []).forEach(function(x) { h += "<th>" + escHtml(x) + "</th>"; });
    h += "</tr></thead><tbody>";
    list.forEach(function(rec) {
        var dealId = rec.Deal_Name && typeof rec.Deal_Name === "object" ? String(rec.Deal_Name.id || "") : "";
        var dn = ddDealName(rec);
        h += "<tr data-id='" + escHtml(rec.id) + "'>";
        h += "<td>" + escHtml(rec.Name || "") + "</td>" + ddEditCell(rec, "ManagementNum");
        h += "<td" + (dealId && !dn ? " data-deal='" + escHtml(dealId) + "'" : "") + ">" + escHtml(dn || (dealId ? "…" : "")) + "</td>";
        h += ddEditCell(rec, "VisitedDateTime") + ddEditCell(rec, "ClosingDay") + ddEditCell(rec, "nousyayoteibi");
        h += ddEditCell(rec, "ServiceStore") + "<td>" + escHtml(personName(rec)) + "</td>" + ddEditCell(rec, "SyaryouCategory");
        if (extraCell) h += extraCell(rec);
        h += "</tr>";
    });
    if (list.length === 0) h += "<tr><td colspan='" + (9 + (extraHead || []).length) + "'>該当する案件はありません</td></tr>";
    h += "</tbody></table>";

    var ov = ddEnsureModal();
    document.getElementById("dd-title").textContent = title;
    document.getElementById("dd-sub").innerHTML = subs.map(escHtml).join("<br>");
    document.getElementById("dd-body").innerHTML = h;
    document.getElementById("dd-body").scrollTop = 0;
    ddBindEdits(document.getElementById("dd-body"));
    ov.style.display = "flex";
    var seq = ++ddSeq;
    ddFillDealNames(list, seq);
    if (ddPicks === null) ddLoadPicks(); // 店舗・車両区分の選択肢は初めて開いたときに1回だけ読む。読めたら今の一覧の欄を選択欄に替える
}

function ddEnsureModal() {
    var ov = document.getElementById("dd-overlay");
    if (ov) return ov;
    ov = document.createElement("div");
    ov.id = "dd-overlay";
    ov.innerHTML = "<div id='dd-box' role='dialog' aria-modal='true'><div id='dd-head'><div id='dd-title'></div><button type='button' id='dd-close' aria-label='閉じる'>×</button></div><div id='dd-sub'></div><div id='dd-body'></div></div>";
    document.body.appendChild(ov);
    ov.addEventListener("click", function(e) { if (e.target === ov) ddClose(); });
    document.getElementById("dd-close").addEventListener("click", ddClose);
    document.getElementById("dd-body").addEventListener("click", function(e) { if (e.target.closest && e.target.closest("td.dd-ed")) return; var tr = e.target.closest ? e.target.closest("tr[data-id]") : null; if (tr) ddOpenRecord(tr.getAttribute("data-id")); });
    return ov;
}
function ddClose() { var ov = document.getElementById("dd-overlay"); if (ov && ov.style.display !== "none") { ov.style.display = "none"; ddSeq++; } }
function ddOpenRecord(id) {
    if (!/^\d+$/.test(String(id || ""))) return;
    // 表を開いたままにするため、新しいタブで開く（liberty 本番の接客 = CustomModule6）。開けなければ同じ画面で開く
    var w = null;
    try { w = window.open("https://crmplus.zoho.jp/crm/org90002333383/tab/CustomModule6/" + id, "_blank"); } catch (e) { w = null; }
    if (w) return;
    try { var pr = ZOHO.CRM.UI.Record.open({ Entity: "Services", RecordID: id }); if (pr && pr.catch) pr.catch(function(e) { console.error("レコードを開けませんでした:", e); }); } catch (e) { console.error("レコードを開けませんでした:", e); }
}

// --- 一覧の中で直せる列（接客 = Services）。保存は接客の画面と同じくワークフローを動かす ---
// recalc: 表の数字に効く項目。保存できたら手元のデータで表を計算し直す（読み込み直しはしない）
var DD_EDIT = {
    ManagementNum:   { kind: "text",     label: "管理番号", recalc: false },
    VisitedDateTime: { kind: "datetime", label: "来店日",   recalc: true },
    ClosingDay:      { kind: "date",     label: "成約日",   recalc: true },
    nousyayoteibi:   { kind: "datetime", label: "納車日",   recalc: true },
    ServiceStore:    { kind: "pick",     label: "店舗",     recalc: true },
    SyaryouCategory: { kind: "pick",     label: "車両区分", recalc: true }
};
var DD_TEXT_MAX = 255;
var DD_NONE = "-None-";
var ddPicks = null;        // 選択肢。null = まだ読んでいない、false = 読めなかった（その列は文字のまま）、{ 項目: [表示名…] }
var ddPicksLoading = false;
var ddCurRecs = {};        // 開いている一覧の id → 案件

// datetime の値を日本時間の "YYYY-MM-DDTHH:mm"（datetime-local の形）に
function ddToJstLocal(v) {
    if (!v) return "";
    var s = String(v), m = s.match(/^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})(?::\d{2}(?:\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?$/);
    if (!m) return "";
    if (!m[3] || m[3] === "+09:00" || m[3] === "+0900") return m[1] + "T" + m[2];
    var d = new Date(s);
    if (isNaN(d.getTime())) return "";
    return new Date(d.getTime() + 9 * 3600 * 1000).toISOString().slice(0, 16);
}
function ddShowVal(field, v) {
    var k = DD_EDIT[field].kind;
    if (v === null || v === undefined) return "";
    if (k === "date") return String(v).slice(0, 10);
    if (k === "datetime") return ddToJstLocal(v);
    return String(v);
}
function ddEditCell(rec, field) {
    var cfg = DD_EDIT[field], v = rec[field], shown = ddShowVal(field, v);
    var attrs = " class='dd-ed-input' data-id='" + escHtml(rec.id) + "' data-f='" + field + "' data-orig='" + escHtml(shown) + "' aria-label='" + escHtml(cfg.label) + "'";
    var ctl;
    if (cfg.kind === "pick") {
        if (!ddPicks || !ddPicks[field]) return "<td data-pf='" + field + "' data-id='" + escHtml(rec.id) + "'>" + escHtml(v || "") + "</td>"; // 選択肢が無い間は文字のまま
        var opts = ddPicks[field].slice();
        if (opts.indexOf(DD_NONE) === -1) opts.unshift(DD_NONE);
        if (shown && opts.indexOf(shown) === -1) opts.unshift(shown); // 今の値が選択肢に無ければ先頭に足す
        var sel = shown || DD_NONE;
        ctl = "<select" + attrs + ">" + opts.map(function(o) { return "<option value='" + escHtml(o) + "'" + (o === sel ? " selected" : "") + ">" + escHtml(o) + "</option>"; }).join("") + "</select>";
    } else {
        var type = cfg.kind === "date" ? "date" : (cfg.kind === "datetime" ? "datetime-local" : "text");
        ctl = "<input type='" + type + "'" + attrs + " value='" + escHtml(shown) + "'>";
    }
    return "<td class='dd-ed'>" + ctl + "<div class='dd-ed-msg'></div></td>";
}
function ddCtlVal(el) {
    var k = DD_EDIT[el.getAttribute("data-f")].kind;
    if (k === "pick") return el.value === DD_NONE ? "" : el.value;
    if (k === "text") return String(el.value || "").trim();
    return el.value || "";
}
function ddRestore(el) {
    var orig = el.getAttribute("data-orig") || "";
    el.value = (el.tagName === "SELECT" && !orig) ? DD_NONE : orig;
}
function ddBindEdits(root) {
    if (!root) return;
    root.querySelectorAll(".dd-ed-input").forEach(function(el) {
        if (el.__bound) return;
        el.__bound = true;
        // 入力欄の操作で行のクリック（レコードを開く）を起こさない
        el.addEventListener("click", function(e) { e.stopPropagation(); });
        el.addEventListener("mousedown", function(e) { e.stopPropagation(); });
        el.addEventListener("keydown", function(e) {
            e.stopPropagation(); // Esc でモーダルを閉じない
            if (e.key === "Enter") { e.preventDefault(); ddSaveEdit(el, true); }
            else if (e.key === "Escape" || e.key === "Esc") { e.preventDefault(); if (!el.__busy) { ddRestore(el); ddEdMsg(el, "", ""); } }
        });
        if (el.tagName === "SELECT") el.addEventListener("change", function() { ddSaveEdit(el, false); });
        else el.addEventListener("blur", function() { ddSaveEdit(el, false); });
    });
}
function ddEdMsg(el, text, kind) {
    var m = el.parentNode ? el.parentNode.querySelector(".dd-ed-msg") : null;
    if (!m) return;
    m.textContent = text; // textContent なので Zoho のエラー文も HTML として解釈されない
    m.className = "dd-ed-msg" + (kind ? " dd-ed-" + kind : "");
}
function ddErrText(x) {
    var d = (x && Array.isArray(x.data) && x.data[0]) ? x.data[0] : (x || {});
    if (typeof d !== "object") return "保存できませんでした: " + String(d);
    var code = d.code || d.status || "", msg = d.message || (x && x.message) || "";
    if (!code && !msg) { try { msg = JSON.stringify(x); } catch (e) { msg = String(x); } }
    return "保存できませんでした: " + [code, msg].filter(function(v) { return v; }).join(" / ");
}
// 保存できた値を手元のデータ（今年・昨年・開いている一覧）に入れる
function ddApplyLocal(id, field, v) {
    [allData || [], lyData || []].forEach(function(arr) { arr.forEach(function(r) { if (r && String(r.id) === id) r[field] = v; }); });
    if (ddCurRecs[id]) ddCurRecs[id][field] = v;
}
async function ddSaveEdit(el, viaEnter) {
    if (!el || el.__busy) return; // 送信中（と確認の最中）の Enter・blur・change は無視
    var field = el.getAttribute("data-f"), cfg = DD_EDIT[field], id = String(el.getAttribute("data-id") || "");
    if (!cfg || !/^\d+$/.test(id)) return;
    var orig = el.getAttribute("data-orig") || "";
    if ((cfg.kind === "date" || cfg.kind === "datetime") && el.validity && el.validity.badInput) { ddEdMsg(el, "日付が途中です。保存していません（Esc で元に戻せます）", "err"); return; }
    var val = ddCtlVal(el);
    if (val === (cfg.kind === "text" ? orig.trim() : orig)) return; // 元と同じなら送らない
    if (cfg.kind === "text" && val.length > DD_TEXT_MAX) { ddEdMsg(el, cfg.label + "は" + DD_TEXT_MAX + "文字までです（いま" + val.length + "文字）。保存していません", "err"); return; }
    el.__busy = true;
    if (val === "") {
        var ok = false;
        try { ok = window.confirm(cfg.label + "を消します。よろしいですか？"); } catch (e) { ok = false; }
        if (!ok) { ddRestore(el); ddEdMsg(el, "", ""); el.__busy = false; return; }
    }
    var send = val === "" ? null : (cfg.kind === "datetime" ? val.slice(0, 16) + ":00+09:00" : val);
    var data = { id: id }; data[field] = send;
    el.disabled = true;
    ddEdMsg(el, "保存しています…", "");
    var res = null, err = null;
    try {
        // ワークフローは止めない（接客の画面で保存したときと同じく、受注データなどへ写すため）
        res = await ZOHO.CRM.API.updateRecord({ Entity: "Services", APIData: data, Trigger: ["workflow"] });
    } catch (e) { err = e; }
    var success = !err && res && Array.isArray(res.data) && res.data[0] && res.data[0].code === "SUCCESS";
    if (success) {
        ddApplyLocal(id, field, send);
        el.setAttribute("data-orig", val);
        ddRestore(el);
        var note = "保存しました";
        if (cfg.recalc) {
            var loadingEl = document.getElementById("loading");
            if (loadingEl && loadingEl.style.display !== "none") note = "保存しました（読み込み中のため、表は読み込みが終わると最新になります）";
            else {
                // 手元のデータで表だけ作り直す。fetchSeq は進めない。モーダルは表の外にあるので消えない
                try { renderAll(); note = "保存しました。表に反映しました（この一覧は開き直すと最新になります）"; }
                catch (e) { console.error("表の再計算エラー:", e); note = "保存しました（表への反映に失敗。画面を開き直してください）"; }
            }
        }
        ddEdMsg(el, note, "ok");
    } else {
        console.error(cfg.label + "の保存エラー:", err || res);
        ddRestore(el);
        ddEdMsg(el, ddErrText(err || res), "err");
    }
    el.disabled = false;
    el.__busy = false;
    if (viaEnter && document.body.contains(el)) { try { el.focus(); } catch (e) {} } // 無効にした間に外れた focus を戻す（続けて Esc で閉じてしまわないように）
}
async function ddLoadPicks() {
    if (ddPicksLoading || ddPicks !== null) return;
    ddPicksLoading = true;
    var got = null;
    try {
        var res = await ZOHO.CRM.META.getFields({ Entity: "Services" });
        var fs = res && (Array.isArray(res.fields) ? res.fields : (Array.isArray(res.data) ? res.data : null));
        if (fs) {
            got = {};
            ["ServiceStore", "SyaryouCategory"].forEach(function(f) {
                var x = fs.filter(function(y) { return y && y.api_name === f; })[0];
                if (x && Array.isArray(x.pick_list_values) && x.pick_list_values.length) got[f] = x.pick_list_values.map(function(p) { return String(p.display_value === null || p.display_value === undefined ? "" : p.display_value); }).filter(function(s) { return s !== ""; });
            });
        }
    } catch (e) { console.error("選択肢の取得エラー:", e); }
    ddPicks = got || false;
    ddPicksLoading = false;
    if (!ddPicks) return;
    // 開いている一覧の文字のままの欄を選択欄に替える
    var body = document.getElementById("dd-body");
    if (!body) return;
    body.querySelectorAll("td[data-pf]").forEach(function(td) {
        var f = td.getAttribute("data-pf"), rec = ddCurRecs[td.getAttribute("data-id")];
        if (!rec || !ddPicks[f]) return;
        var tmp = document.createElement("tbody");
        tmp.innerHTML = "<tr>" + ddEditCell(rec, f) + "</tr>";
        td.parentNode.replaceChild(tmp.firstChild.firstChild, td);
    });
    ddBindEdits(body);
}

// お客様名（商談名）が COQL で返らなかったときだけ、商談を 50 件ずつ読んで埋める（読み取りだけ）
async function ddQueryDeals(ids) {
    var forms = [ids.join(", "), ids.map(function(id) { return "'" + id + "'"; }).join(", ")];
    for (var i = 0; i < forms.length; i++) {
        try {
            var res = await ZOHO.CRM.API.coql({ "select_query": "select Deal_Name from Deals where id in (" + forms[i] + ") limit 0, 50" });
            if (res && Array.isArray(res.data)) return res.data;
            if (res && typeof res === "object" && !res.code && !res.message && res.status !== "error") return []; // 0件
        } catch (e) { if (i === forms.length - 1) console.error("お客様名の取得エラー:", e); }
    }
    return null;
}
async function ddFillDealNames(recs, seq) {
    var ids = [];
    recs.forEach(function(rec) { var d = rec.Deal_Name; if (d && typeof d === "object" && !d.name && d.id) { var id = String(d.id); if (/^\d+$/.test(id) && !(id in dealNameMap) && ids.indexOf(id) === -1) ids.push(id); } });
    var failed = {};
    for (var i = 0; i < ids.length; i += 50) {
        var chunk = ids.slice(i, i + 50);
        var data = await ddQueryDeals(chunk);
        if (data) { data.forEach(function(x) { if (x && x.id) dealNameMap[String(x.id)] = x.Deal_Name || ""; }); chunk.forEach(function(id) { if (!(id in dealNameMap)) dealNameMap[id] = ""; }); }
        else chunk.forEach(function(id) { failed[id] = true; });
        if (seq !== ddSeq) return;
    }
    if (seq !== ddSeq) return;
    document.querySelectorAll("#dd-body td[data-deal]").forEach(function(td) { var id = td.getAttribute("data-deal"); td.textContent = failed[id] ? "-" : (dealNameMap[id] || "-"); });
}

document.addEventListener("click", function(e) {
    if (!e.target || !e.target.closest) return;
    var td = e.target.closest("td.dd-cell");
    if (!td) return;
    var cat = e.target.closest(".val-kei") ? "k" : (e.target.closest(".val-fu") ? "f" : "t");
    ddOpenFromCell(td.getAttribute("data-dd"), cat);
});
document.addEventListener("keydown", function(e) { if (e.key === "Escape" || e.key === "Esc") ddClose(); });
