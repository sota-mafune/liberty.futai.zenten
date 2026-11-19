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
var SELECT_FIELDS = "ClosingDay, VisitedDateTime, nousyayoteibi, SyaryouCategory, FOrR, Seated, Option1, ServiceStore, ServicePerson, ServicePerson.Name, cancel, HanbaiCategory, Option2, Option3, Option15, Option4, Option5, Option16, Option7, Option8, Option9, Option10, Option6, BackCamera, Option17, TradeinCar, PaymentCategory, arari16, arari17, Option14, arari21, arari22, arari23, arari24, arari25";

// 集計の期間（renderAll の中で使う）
var curRange = null;

if (typeof ZOHO !== "undefined" && ZOHO.embeddedApp) {
    ZOHO.embeddedApp.init().then(function() {
        initMonthSelector();
        setInitialDates();
        fetchByCOQL();
    });
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
async function fetchServices(where, seq) {
    var out = [], offset = 0;
    while (true) {
        var res = await ZOHO.CRM.API.coql({ "select_query": "select " + SELECT_FIELDS + " from Services where " + where + " limit " + offset + ", 200" });
        if (isStale(seq)) return null;
        if (res && Array.isArray(res.data)) {
            out = out.concat(res.data);
            if (res.info && res.info.more_records) { offset += 200; continue; }
            break;
        }
        if (res && typeof res === "object" && (res.code || res.status === "error" || res.message)) throw new Error(res.code || res.message || "COQL error");
        break; // 0件
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
    var cur = [], ly = [], lyOk = false;
    try {
        cur = await fetchServices(where, seq);
        if (cur === null || isStale(seq)) return;
        // 昨年分は前年の期間が CRM のデータの開始日以降のときだけ読む
        if (p.lyFetchStart && p.lyFetchStart >= CRM_DATA_START) {
            loadingEl.innerHTML += "<br>▶ 昨年の実績を取得中...";
            var lyWhere = "((ClosingDay between '" + p.lyFetchStart + "' and '" + p.lyFetchEnd + "') or (VisitedDateTime between '" + p.lyFetchStart + "T00:00:00+09:00' and '" + p.lyFetchEnd + "T23:59:59+09:00'))";
            ly = await fetchServices(lyWhere, seq);
            if (ly === null || isStale(seq)) return;
            lyOk = true;
        }
    } catch (e) {
        if (isStale(seq)) return;
        console.error("COQL 取得エラー:", e);
        // 前の月の集計を残さない（フィルタを変えても前の表が出ないように）
        allData = []; lyData = []; lyAvailable = false; budgetDataGlobal = null;
        storeStatsMaster = {}; staffStatsMaster = {}; storeGroupMap = {}; staffStoreMap = {}; loadFailed = true;
        loadingEl.innerHTML += "<br><span style='color:#c00000;'>▶ 取得に失敗しました（" + String((e && e.message) || e).replace(/[<>&]/g, "") + "）。期間を選び直すか、画面を開き直してください。</span>";
        updateSelector('group-selector', [], '全グループ表示'); updateSelector('store-selector', [], '全店舗表示');
        ["group-table-container", "store-table-container", "staff-table-container"].forEach(function(id) { document.getElementById(id).innerHTML = ""; });
        return;
    }

    await resolveMastersNames(cur.concat(ly));
    if (isStale(seq)) return;
    var budget = await fetchAnalyticsBudgets();
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

async function fetchAnalyticsBudgets() {
    try {
        const res = await ZOHO.CRM.FUNCTIONS.execute("get_dashboard_budgets", { arguments: JSON.stringify({}) });
        if(res && res.code && res.code.toLowerCase() === "success" && res.details && res.details.output) {
            return JSON.parse(res.details.output);
        }
    } catch(e) { console.error("❌ 予算データ取得エラー:", e); }
    return null;
}

var STAT_BASES = ["j", "v_n", "sho", "ab", "jk", "rv", "rj", "tot_v", "n", "m", "c", "o2", "o3", "pk", "ct", "up", "tp", "ic", "rst", "ni", "nu", "hp", "fl", "aq", "tr", "ln", "l84", "r69", "r59", "r49", "r39", "r29", "low", "zn", "ar21", "ar22", "ar23", "ar24", "ar25", "ar_cnt", "del_cnt", "del_ar21", "del_ar22", "del_ar23", "del_ar24", "del_ar25", "ly_j", "ly_jt", "ly_vt"];
// budget_rows: 月次予算の行数（0 なら「未登録」）、daily_rows: その月の日次予算の行数
function createStats() { var s = { budget_n:0, budget_ar:0, budget_j:0, budget_current:0, budget_rows:0, daily_rows:0, g_sls:0 }; STAT_BASES.forEach(function(b) { s[b + "_k"] = 0; s[b + "_f"] = 0; }); return s; }

// コーティング・室内コート・防錆コート・ナビアップ・希望ナンバーは「空・✖・×・-None-」以外ならすべて獲得
function isGot(v) { if (v === null || v === undefined) return false; v = String(v).trim(); return ["", "✖", "×", "-None-"].indexOf(v) === -1; }
function catOf(rec) { return (rec.SyaryouCategory === "軽" || rec.SyaryouCategory === "軽自動車") ? "k" : "f"; }
function isCancelled(rec) { return rec.cancel === true || rec.cancel === "true"; }

function aggregate(s, rec) {
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
}

// 昨年: ly_j = 前年の同じ月の成約台数、ly_jt = 前年の開始日〜min(終了日,今日)の1年前までの成約台数、ly_vt = 同じ期間の新規接客数
function aggregateLY(s, rec) {
    var c = catOf(rec);
    var vD = (rec.VisitedDateTime || "").split('T')[0], cD = (rec.ClosingDay || ""), isCancel = isCancelled(rec);
    var R = curRange;
    if (cD && !isCancel) {
        if (cD >= R.lyMonthStart && cD <= R.lyMonthEnd) s["ly_j_"+c]++;
        if (cD >= R.lyStart && cD <= R.lyCutEnd) s["ly_jt_"+c]++;
    }
    if (vD && vD >= R.lyStart && vD <= R.lyCutEnd && rec.FOrR === "初回") s["ly_vt_"+c]++;
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

    if(budgetDataGlobal) {
        var month = curRange.month, today = curRange.today, todayMonth = today.slice(0, 7);

        if(budgetDataGlobal.sales_budget) {
            budgetDataGlobal.sales_budget.forEach(b => {
                if(normDate(b["月"]).slice(0, 7) !== month) return;
                var val_j = parseInt(b["成約台数予算"]) || 0, val_n = parseInt(b["納車予算"]) || 0, val_ar = parseInt(b["粗利予算"]) || 0;
                bucketsFor(b["店舗"] || "未所属", staffKey(b["担当者"]), "budget").forEach(s => { s.budget_j += val_j; s.budget_n += val_n; s.budget_ar += val_ar; s.budget_rows++; });
            });
        }
        if(budgetDataGlobal.daily_budget) {
            budgetDataGlobal.daily_budget.forEach(b => {
                var d = normDate(b["日"]);
                if(d.slice(0, 7) !== month) return;
                // 現時点予算: 今月は今日まで、過去の月は月の合計、未来の月は 0
                var include = month < todayMonth ? true : (month === todayMonth ? d <= today : false);
                var val_cur = include ? (parseInt(b["成約台数予算"]) || 0) : 0;
                bucketsFor(b["店舗"] || "未所属", null, "budget").forEach(s => { s.budget_current += val_cur; s.daily_rows++; });
            });
        }
    }

    // 昨年は列を作ったあとに足す（昨年だけの店舗・担当者の列は作らない）
    if (lyAvailable) (lyData || []).forEach(r => { bucketsFor(r.ServiceStore || "未所属", staffKey(personName(r)), "ly").forEach(s => aggregateLY(s, r)); });

    // 500ms 後に読み込み表示を消す。その間に新しい読み込みが始まっていたら消さない
    setTimeout(() => { if (seq === undefined || !isStale(seq)) loadingEl.style.display = 'none'; }, 500);
    updateSelector('group-selector', orderKeys(Array.from(groupSet), "group"), '全グループ表示');
    updateSelector('store-selector', orderKeys(Array.from(storeSet), "store"), '全店舗表示');
    document.getElementById("group-table-container").innerHTML = buildTable(gS, "グループ名", totalS, "group");
    filterStoreByGroup(); filterStaffByStore();
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
    var h = "<table><thead><tr>";
    h += "<th class='sticky-col-item shop-header' style='position: sticky !important; z-index: 300 !important; top: 0; left: 0;'>" + (title || "KPI項目") + "</th>";
    h += "<th class='sticky-col-total shop-header' style='position: sticky !important; z-index: 290 !important; top: 0; left: 170px;'>合計</th>";
    for(var i=0; i<keys.length; i++) { h += "<th class='shop-header' style='position: sticky !important; z-index: 150 !important; top: 0;'>" + keys[i] + "</th>"; }
    h += "</tr></thead><tbody>";

    // d2: 2段目（合計）の分母（キーを足し合わせる）。無ければ d と同じ
    const rowDef = [
        { sec: "予算・目標" },
        { lbl: "予算", m: "budget_j", type: "total_only", flag: "budget_rows", cls: "#ffffff" },
        { lbl: "目標", type: "total_dash", cls: "#ffffff" },
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

    for(var j=0; j<rowDef.length; j++){
        var r = rowDef[j];
        if(r.sec) {
            h += "<tr><td class='sticky-col-item section-row' style='position: sticky !important; z-index: 180 !important; left: 0; border-right: none;'>" + r.sec + "</td><td class='sticky-col-total section-row' style='position: sticky !important; z-index: 170 !important; left: 170px; border-left: none;'></td>";
            if (keys.length > 0) h += "<td colspan='" + keys.length + "' class='section-row' style='position: static; z-index: 1; border-left: none;'></td>";
            h += "</tr>";
        } else {
            h += "<tr><td class='sticky-col-item' style='background-color:"+(r.cls||"#fff")+"'>"+r.lbl+"</td>";
            h += renderCell(totalS, r, true);
            for(var k=0; k<keys.length; k++) h += renderCell(sum[keys[k]], r, false);
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
// 総粗利の軽/普の段は幅に収まるよう万円単位
function man(v) { return Math.round(v / 10000).toLocaleString() + "万"; }

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
        tVal = (r.flag && !s[r.flag]) ? "未登録" : (s[r.m] || 0).toLocaleString();
        return singleCell(c, cellStyle, tVal, " color:#444;");
    }
    if(r.type === "total_dash") return singleCell(c, cellStyle, "-", " color:#444;");

    // 2. 達成率（分母の予算が未登録・0 なら「-」）
    if(r.type === "total_ratio" || r.type === "del_total_ratio") {
        var act = (s[r.n+"_k"] || 0) + (s[r.n+"_f"] || 0);
        var bud = s[r.d] || 0;
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
            kVal = man(kS); fVal = man(fS);
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
function filterStoreByGroup() { if (loadFailed) { document.getElementById("store-table-container").innerHTML = ""; return; } var g = document.getElementById('group-selector').value, f = {}, t = createStats(); Object.keys(storeStatsMaster).forEach(function(st){ if (g === "all" || storeGroupMap[st] === g) { f[st] = storeStatsMaster[st]; Object.keys(t).forEach(function(k){ if(typeof t[k] === 'number') t[k] += storeStatsMaster[st][k]; }); } }); document.getElementById("store-table-container").innerHTML = buildTable(f, "店舗名", t, "store"); }
function filterStaffByStore() { if (loadFailed) { document.getElementById("staff-table-container").innerHTML = ""; return; } var s = document.getElementById('store-selector').value, f = {}, t = createStats(); Object.keys(staffStatsMaster).forEach(function(n){ if (s === "all" || staffStoreMap[n] === s) { f[n] = staffStatsMaster[n]; Object.keys(t).forEach(function(k){ if(typeof t[k] === 'number') t[k] += staffStatsMaster[n][k]; }); } }); document.getElementById("staff-table-container").innerHTML = buildTable(f, "担当者名", t, "staff"); }
