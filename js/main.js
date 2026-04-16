var allData = [];
var budgetDataGlobal = null; 
var staffStatsMaster = {}, storeStatsMaster = {}, storeGroupMap = {}, staffStoreMap = {};
var personMap = {};

var storeToGroup = { "神戸店":"兵四", "久米窪田店":"兵四", "高知高須店":"兵四", "北久米店":"兵四", "尼崎店":"兵四", "高槻店":"大阪", "八尾店":"大阪", "堺大泉緑地前店":"大阪", "松原天美店":"大阪", "貝塚店":"大阪", "大津店":"滋三", "栗東店":"滋三", "彦根店":"滋三", "津店":"滋三", "松阪店":"滋三", "鯖江店":"滋三", "久御山店":"京奈", "171店":"京奈", "精華店":"京奈", "西大和店":"京奈", "橿原店":"京奈", "熊本インター店":"旧Dj", "長田店":"旧Dj", "outlet店":"旧Dj", "舞鶴店":"旧Dj", "福知山店":"旧Dj", "加古川店":"旧Dj", "BYD滋賀":"未所属" };

ZOHO.embeddedApp.init().then(function() {
    initMonthSelector();
    setInitialDates();
    fetchByCOQL();
});

function initMonthSelector() { var sel = document.getElementById('month-selector'); var now = new Date(); var curMonth = now.getFullYear() + "-" + ("0" + (now.getMonth() + 1)).slice(-2); for(var y=2025; y<=2026; y++) { for(var m=1; m<=12; m++) { var v = y + "-" + (m < 10 ? "0" + m : m); var opt = document.createElement('option'); opt.value = v; opt.text = y + "/" + m; if(v === curMonth) opt.selected = true; sel.add(opt); } } }
function setInitialDates() { var now = new Date(); var y = now.getFullYear(), m = now.getMonth() + 1, last = new Date(y, m, 0).getDate(), mStr = ("0" + m).slice(-2); document.getElementById('start-date').value = y + "-" + mStr + "-01"; document.getElementById('end-date').value = y + "-" + mStr + "-" + last; }
function syncMonthToCalendar() { var v = document.getElementById('month-selector').value, p = v.split("-").map(Number), l = new Date(p[0], p[1], 0).getDate(); document.getElementById('start-date').value = v + "-01"; document.getElementById('end-date').value = v + "-" + l; fetchByCOQL(); }

async function fetchByCOQL() {
    allData = []; budgetDataGlobal = null;
    var loadingEl = document.getElementById('loading');
    loadingEl.style.display = 'block';
    loadingEl.style.textAlign = 'left';
    loadingEl.innerHTML = "<b>【データ読み込み状況】</b><br>▶ 実績データを取得中...";
    
    const start = document.getElementById('start-date').value, end = document.getElementById('end-date').value;
    if(!start || !end) return;

    let page = 1, hasMore = true;
    while(hasMore) {
        const offset = (page - 1) * 200;
        // ★ ServicePerson.Name を復活させ、もし nousyayoteibi が原因ならここでエラーが出るようにします
        const coql = { "select_query": "select ClosingDay, VisitedDateTime, nousyayoteibi, SyaryouCategory, FOrR, Seated, Option1, ServiceStore, ServicePerson.Name, cancel, HanbaiCategory, Option2, Option3, Option15, Option4, Option5, Option16, Option7, Option8, Option9, Option10, Option6, BackCamera, Option17, TradeinCar, PaymentCategory, arari16, arari17, Option14, arari21, arari22, arari23, arari24, arari25 from Services where ((ClosingDay between '" + start + "' and '" + end + "') or (VisitedDateTime between '" + start + "T00:00:00+09:00' and '" + end + "T23:59:59+09:00') or (nousyayoteibi between '" + start + "' and '" + end + "')) limit " + offset + ", 200" };
        try { 
            const res = await ZOHO.CRM.API.coql(coql); 
            if (res.data) { 
                allData = allData.concat(res.data); 
                if (res.info && res.info.more_records) page++; else hasMore = false; 
            } else {
                console.error("COQLエラーレスポンス:", res); // エラーならF12で確認可能
                hasMore = false; 
            }
        } catch (e) { 
            console.error("COQL例外発生:", e);
            hasMore = false; 
        }
    } 

    await fetchAnalyticsBudgets();
    renderAll();
}

async function fetchAnalyticsBudgets() {
    try {
        const res = await ZOHO.CRM.FUNCTIONS.execute("get_dashboard_budgets", { arguments: JSON.stringify({}) });
        if(res && res.code && res.code.toLowerCase() === "success" && res.details && res.details.output) {
            budgetDataGlobal = JSON.parse(res.details.output);
            console.log("✅ 予算データの受信に成功しました！", budgetDataGlobal);
        } else {
            console.warn("⚠️ 予算取得：予期せぬレスポンス", res);
        }
    } catch(e) {
        console.error("❌ 予算データ取得エラー:", e);
    }
}

// ★ 追加：納車予算(budget_del)、粗利予算(budget_arari)、納車ベースの集計箱(del_xxx)を追加
function createStats() { return { budget_j:0, budget_current:0, budget_del:0, budget_arari:0, j_k:0, j_f:0, v_n_k:0, v_n_f:0, sho_k:0, sho_f:0, ab_k:0, ab_f:0, jk_k:0, jk_f:0, rv_k:0, rv_f:0, rj_k:0, rj_f:0, tot_v_k:0, tot_v_f:0, n_k:0, n_f:0, m_k:0, m_f:0, c_k:0, c_f:0, o2_k:0, o2_f:0, o3_k:0, o3_f:0, pk_k:0, pk_f:0, ct_k:0, ct_f:0, up_k:0, up_f:0, tp_k:0, tp_f:0, ic_k:0, ic_f:0, rst_k:0, rst_f:0, ni_k:0, ni_f:0, nu_k:0, nu_f:0, hp_k:0, hp_f:0, fl_k:0, fl_f:0, aq_k:0, aq_f:0, tr_k:0, tr_f:0, ln_k:0, ln_f:0, l84_k:0, l84_f:0, r69_k:0, r69_f:0, r59_k:0, r59_f:0, r49_k:0, r49_f:0, r39_k:0, r39_f:0, r29_k:0, r29_f:0, low_k:0, low_f:0, zn_k:0, zn_f:0, ar21_k:0, ar21_f:0, ar22_k:0, ar22_f:0, ar23_k:0, ar23_f:0, ar24_k:0, ar24_f:0, ar25_k:0, ar25_f:0, ar_cnt_k:0, ar_cnt_f:0, del_cnt_k:0, del_cnt_f:0, del_ar21_k:0, del_ar21_f:0, del_ar22_k:0, del_ar22_f:0, del_ar23_k:0, del_ar23_f:0, del_ar24_k:0, del_ar24_f:0, del_ar25_k:0, del_ar25_f:0 }; }

function aggregate(s, rec) {
    var c = (rec.SyaryouCategory === "軽" || rec.SyaryouCategory === "軽自動車") ? "k" : "f";
    var vD = (rec.VisitedDateTime || "").split('T')[0], cD = (rec.ClosingDay || ""), nD = (rec.nousyayoteibi || "").split('T')[0], isCancel = (rec.cancel === true || rec.cancel === "true");
    const st = document.getElementById('start-date').value, ed = document.getElementById('end-date').value;
    
    // 来店日ベース
    if (vD && vD >= st && vD <= ed) { s["tot_v_"+c]++; if (rec.FOrR === "初回") s["v_n_"+c]++; if (rec.Seated === "〇") s["sho_"+c]++; if (rec.Option1 === "〇") s["ab_"+c]++; if (rec.FOrR === "初回" && cD === vD && !isCancel) s["jk_"+c]++; if (rec.FOrR === "再来") s["rv_"+c]++; }
    
    // 計上日ベース（受注時想定など）
    if (cD && cD >= st && cD <= ed && !isCancel) { s["j_"+c]++; if (rec.FOrR === "再来") s["rj_"+c]++; if (rec.HanbaiCategory === "新") s["n_"+c]++; if (rec.HanbaiCategory === "未") s["m_"+c]++; if (rec.HanbaiCategory === "中") s["c_"+c]++; if (rec.Option2 === "〇") s["o2_"+c]++; if (rec.Option3 === "〇") s["o3_"+c]++; if (["新車ﾊﾟｯｸ","未使用ﾊﾟｯｸ","中ｴｺﾊﾟｯｸ","中ｽﾀﾊﾟｯｸ","中ｱﾌﾟﾊﾟｯｸ"].indexOf(rec.Option15) !== -1) s["pk_"+c]++; if (rec.Option4 === "〇") s["ct_"+c]++; if (rec.Option5 === "〇") s["up_"+c]++; if (rec.Option16 === "〇") s["tp_"+c]++; if (rec.Option7 === "〇") s["ic_"+c]++; if (rec.Option8 === "〇") s["rst_"+c]++; if (rec.Option9 === "〇") s["ni_"+c]++; if (rec.Option10 === "〇") s["nu_"+c]++; if (rec.Option6 === "〇") s["hp_"+c]++; if (rec.BackCamera === "〇") s["fl_"+c]++; if (rec.Option17 === "〇") s["aq_"+c]++; if (["✖","✕","×","","null","-","無","なし"].indexOf(rec.TradeinCar || "") === -1) s["tr_"+c]++; if ((rec.PaymentCategory || "").indexOf("ローン") !== -1) { s["ln_"+c]++; var r = parseFloat(rec.arari17)||0, ct = parseInt(rec.arari16)||0; if (ct >= 84) s["l84_"+c]++; if (r >= 6 && r < 7) s["r69_"+c]++; else if (r >= 5 && r < 6) s["r59_"+c]++; else if (r >= 4 && r < 5) s["r49_"+c]++; else if (r >= 3 && r < 4) s["r39_"+c]++; else if (r >= 2.9 && r < 3) s["r29_"+c]++; else if (r > 0 && r < 2.9) s["low_"+c]++; } if (rec.Option14 === "〇") s["zn_"+c]++; s["ar21_"+c] += (parseFloat(rec.arari21)||0); s["ar22_"+c] += (parseFloat(rec.arari22)||0); s["ar23_"+c] += (parseFloat(rec.arari23)||0); s["ar24_"+c] += (parseFloat(rec.arari24)||0); s["ar25_"+c] += (parseFloat(rec.arari25)||0); s["ar_cnt_"+c]++; }
    
    // ★ 納車日ベース（納車着地粗利予測）
    if (nD && nD >= st && nD <= ed && !isCancel) {
        s["del_cnt_"+c]++;
        s["del_ar21_"+c] += (parseFloat(rec.arari21)||0);
        s["del_ar22_"+c] += (parseFloat(rec.arari22)||0);
        s["del_ar23_"+c] += (parseFloat(rec.arari23)||0);
        s["del_ar24_"+c] += (parseFloat(rec.arari24)||0);
        s["del_ar25_"+c] += (parseFloat(rec.arari25)||0);
    }
}

function renderAll() {
    document.getElementById('loading').style.display = 'none';
    var gS = {}, totalS = createStats(); staffStatsMaster = {}; storeStatsMaster = {}; var groupSet = new Set(), storeSet = new Set();
    
    // 1. 実績の集計
    if(allData && allData.length > 0) {
        allData.forEach(r => {
            var st = r.ServiceStore || "未所属", gr = storeToGroup[st] || "未所属";
            var pr = (r.ServicePerson && r.ServicePerson.Name) ? r.ServicePerson.Name : "未設定";
            if(!gS[gr]) gS[gr] = createStats(); if(!storeStatsMaster[st]) storeStatsMaster[st] = createStats(); if(!staffStatsMaster[pr]) staffStatsMaster[pr] = createStats();
            storeGroupMap[st] = gr; staffStoreMap[pr] = st; groupSet.add(gr); storeSet.add(st);
            [gS[gr], storeStatsMaster[st], staffStatsMaster[pr], totalS].forEach(s => aggregate(s, r));
        });
    }

    // 2. 予算の集計
    if(budgetDataGlobal) {
        var selectedMonth = document.getElementById('month-selector').value; 
        var targetDateStr = selectedMonth.replace("-", "/"); 
        var endDateStr = document.getElementById('end-date').value.replace(/-/g, "/"); 
        if(budgetDataGlobal.sales_budget) {
            budgetDataGlobal.sales_budget.forEach(b => {
                var d = b["月"] || "";
                if(d.includes(targetDateStr) || d.includes(selectedMonth)) {
                    var st = b["店舗"], pr = b["担当者"], gr = storeToGroup[st] || "未所属";
                    if(!gS[gr]) gS[gr] = createStats(); if(!storeStatsMaster[st]) storeStatsMaster[st] = createStats(); if(!staffStatsMaster[pr]) staffStatsMaster[pr] = createStats();
                    storeGroupMap[st] = gr; staffStoreMap[pr] = st; groupSet.add(gr); storeSet.add(st);
                    gS[gr].budget_j += parseInt(b["成約台数予算"])||0; storeStatsMaster[st].budget_j += parseInt(b["成約台数予算"])||0; staffStatsMaster[pr].budget_j += parseInt(b["成約台数予算"])||0; totalS.budget_j += parseInt(b["成約台数予算"])||0;
                    gS[gr].budget_del += parseInt(b["納車予算"])||0; storeStatsMaster[st].budget_del += parseInt(b["納車予算"])||0; staffStatsMaster[pr].budget_del += parseInt(b["納車予算"])||0; totalS.budget_del += parseInt(b["納車予算"])||0;
                    var arBudget = parseInt((b["粗利予算"]||"0").toString().replace(/[^0-9]/g, ''))||0;
                    gS[gr].budget_arari += arBudget; storeStatsMaster[st].budget_arari += arBudget; staffStatsMaster[pr].budget_arari += arBudget; totalS.budget_arari += arBudget;
                }
            });
        }
        if(budgetDataGlobal.daily_budget) {
            budgetDataGlobal.daily_budget.forEach(b => {
                var d = (b["日"] || "").replace(/-/g, "/"); 
                if(d.startsWith(targetDateStr) && d <= endDateStr) {
                    var st = b["店舗"], gr = storeToGroup[st] || "未所属";
                    if(!gS[gr]) gS[gr] = createStats(); if(!storeStatsMaster[st]) storeStatsMaster[st] = createStats();
                    storeGroupMap[st] = gr; groupSet.add(gr); storeSet.add(st);
                    var dVal = parseInt(b["成約台数予算"])||0;
                    gS[gr].budget_current += dVal; storeStatsMaster[st].budget_current += dVal; totalS.budget_current += dVal;
                }
            });
        }
    }

    updateSelector('group-selector', groupSet, '全グループ表示'); updateSelector('store-selector', storeSet, '全店舗表示');
    document.getElementById("group-table-container").innerHTML = buildTable(gS, "グループ名", totalS);
    filterStoreByGroup(); filterStaffByStore();
}

function updateSelector(id, set, def) { var sel = document.getElementById(id); sel.innerHTML = '<option value="all">' + def + '</option>'; Array.from(set).sort().forEach(v => { var o = document.createElement('option'); o.value = o.text = v; sel.add(o); }); }

function buildTable(sum, title, totalS) {
    var keys = Object.keys(sum).sort();
    var h = "<table><thead><tr>";
    h += "<th class='sticky-col-item shop-header' style='position: sticky !important; z-index: 300 !important; top: 0; left: 0;'>" + (title || "KPI項目") + "</th>";
    h += "<th class='sticky-col-total shop-header' style='position: sticky !important; z-index: 290 !important; top: 0; left: 170px;'>合計</th>";
    for(var i=0; i<keys.length; i++) { h += "<th class='shop-header' style='position: sticky !important; z-index: 150 !important; top: 0;'>" + keys[i] + "</th>"; }
    h += "</tr></thead><tbody>";

    const rowDef = [
        { sec: "予算・目標" }, 
        { lbl: "予算", m: "budget_j", type: "total_only", cls: "#ffffff" },
        { lbl: "昨年実績", m: "empty", cls: "#ffffff" }, 
        { lbl: "現時点予算", m: "budget_current", type: "total_only", cls: "#ffffff" },
        { sec: "基本実績" }, 
        { lbl: "実績", m: "j", cls: "#ffe599" }, 
        { lbl: "達成率", type: "total_ratio", n: "j", d: "budget_j", cls: "#ffffff" },
        { lbl: "昨年実績(当日)", m: "empty", cls: "#d9d2e9" }, { lbl: "昨年対比", m: "empty", cls: "#d9d2e9" },
        { lbl: "新規接客数", m: "v_n", cls: "#ffffff" }, { lbl: "昨年接客数(当日)", m: "empty", cls: "#d9d2e9" }, { lbl: "商談件数", m: "sho", cls: "#ffffff" }, { lbl: "商談率", type: "ratio", n: "sho", d: "v_n", cls: "#ffe599" },
        { lbl: "AB数", m: "ab", cls: "#ffffff" }, { lbl: "AB率", type: "ratio", n: "ab", d: "sho", cls: "#ffffff" }, { lbl: "即決成約", m: "jk", cls: "#ffffff" }, { lbl: "即決率", type: "ratio", n: "jk", d: "v_n", cls: "#ffffff" },
        { lbl: "成約率", type: "ratio", n: "j", d: "v_n", cls: "#ffe599", redText: true }, { lbl: "昨年成約率", m: "empty", cls: "#d9d2e9" },
        { lbl: "再来接客数", m: "rv", cls: "#ffffff" }, { lbl: "再来店率", type: "custom_ratio", n: "rv", d_sub: ["v_n", "jk"], cls: "#ffffff" },
        { lbl: "再来成約", m: "rj", cls: "#ffffff" }, { lbl: "再来成約率", type: "custom_ratio", n: "rj", d_sub: ["v_n", "jk"], cls: "#ffffff" },
        { lbl: "総接客数", m: "tot_v", cls: "#ffffff" }, { lbl: "総成約率", type: "ratio", n: "j", d: "tot_v", cls: "#ffffff" },
        { lbl: "追客可能数", type: "diff", n: "v_n", d: "j", cls: "#f4cccc" }, { lbl: "決着済", m: "empty", cls: "#f4cccc" }, { lbl: "決着率", m: "empty", cls: "#f4cccc" },
        { sec: "販売区分" }, { lbl: "新車", m: "n", cls: "#b6d7a8" }, { lbl: "獲得率", type: "ratio", n: "n", d: "j", cls: "#ffffff" }, { lbl: "未使用車", m: "m", cls: "#b6d7a8" }, { lbl: "獲得率", type: "ratio", n: "m", d: "j", cls: "#ffffff" }, { lbl: "中古車", m: "c", cls: "#b6d7a8" }, { lbl: "獲得率", type: "ratio", n: "c", d: "j", cls: "#ffffff" },
        { sec: "付帯品個別" }, { lbl: "10年保証", m: "o2", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "o2", d: "j", cls: "#ffffff" }, { lbl: "中古保証", m: "o3", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "o3", d: "j", cls: "#ffffff" }, { lbl: "パック", m: "pk", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "pk", d: "j", cls: "#ffffff" }, { lbl: "コーティング", m: "ct", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "ct", d: "j", cls: "#ffffff" }, { lbl: "UPグレード", m: "up", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "up", d: "j", cls: "#ffffff" }, { lbl: "T-プレミアム", m: "tp", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "tp", d: "j", cls: "#ffffff" }, { lbl: "室内コート", m: "ic", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "ic", d: "j", cls: "#ffffff" }, { lbl: "防錆コート", m: "rst", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "rst", d: "j", cls: "#ffffff" }, { lbl: "ナビ取付", m: "ni", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "ni", d: "j", cls: "#ffffff" }, { lbl: "ナビアップ", m: "nu", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "nu", d: "j", cls: "#ffffff" }, { lbl: "希望ナンバー", m: "hp", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "hp", d: "j", cls: "#ffffff" }, { lbl: "フィルム", m: "fl", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "fl", d: "j", cls: "#ffffff" }, { lbl: "アクアペル", m: "aq", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "aq", d: "j", cls: "#ffffff" }, { lbl: "下取り", m: "tr", cls: "#a2c4c9" }, { lbl: "獲得率", type: "ratio", n: "tr", d: "j", cls: "#ffffff" },
        { sec: "総付帯実績" }, { lbl: "車両総付帯", type: "sum", items: ["ct","up","tp","ic","rst","nu","fl","aq"], cls: "#a2c4c9" }, { lbl: "獲得率", type: "sum_ratio", items: ["ct","up","tp","ic","rst","nu","fl","aq"], d: "j", cls: "#ffffff" }, { lbl: "周辺総付帯", type: "sum", items: ["n","tr","ln","r69"], cls: "#a2c4c9" }, { lbl: "獲得率", type: "sum_ratio", items: ["n","tr","ln","r69"], d: "j", cls: "#ffffff" },
        { sec: "ローン実績" }, { lbl: "ローン", m: "ln", cls: "#a4c2f4" }, { lbl: "ローン獲得率", type: "ratio", n: "ln", d: "j", cls: "#ffffff" }, { lbl: "84回以上", m: "l84", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "l84", d: "j", cls: "#ffffff" }, { lbl: "6.90%", m: "r69", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r69", d: "j", cls: "#ffffff" }, { lbl: "5.90%", m: "r59", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r59", d: "j", cls: "#ffffff" }, { lbl: "4.90%", m: "r49", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r49", d: "j", cls: "#ffffff" }, { lbl: "3.90%", m: "r39", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r39", d: "j", cls: "#ffffff" }, { lbl: "2.90%", m: "r29", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "r29", d: "j", cls: "#ffffff" }, { lbl: "低金利", m: "low", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "low", d: "j", cls: "#ffffff" }, { lbl: "残価設定", m: "zn", cls: "#a4c2f4" }, { lbl: "獲得率", type: "ratio", n: "zn", d: "j", cls: "#ffffff" },
        { sec: "受注時想定" }, { lbl: "受注台数", type: "arari_val", val: "j", cls: "#ead1dc" }, { lbl: "@車両粗利", type: "arari_avg", val: "ar21", cls: "#ead1dc" }, { lbl: "@ローンBK", type: "arari_avg", val: "ar23", cls: "#ead1dc" }, { lbl: "@下取粗利", type: "arari_avg", val: "ar22", cls: "#ead1dc" }, { lbl: "@全部割(保証抜き)", type: "arari_avg", val: "ar24", cls: "#ead1dc" }, { lbl: "@全部割(保証込み)", type: "arari_avg", val: "ar25", cls: "#ead1dc" }, { lbl: "総粗利", type: "arari_sum", val: "ar25", cls: "#ead1dc" },
        // ★ 納車着地粗利予測 セクション追加
        { sec: "納車着地粗利予測 <span style='font-size:10px; font-weight:normal; color:#d32f2f;'>(※下取粗利は概算)</span>" },
        { lbl: "納車台数", type: "del_arari_val", cls: "#d9ead3" },
        { lbl: "@車両粗利", type: "del_arari_avg", val: "del_ar21", cls: "#d9ead3" },
        { lbl: "@ローンBK", type: "del_arari_avg", val: "del_ar23", cls: "#d9ead3" },
        { lbl: "@下取粗利", type: "del_arari_avg", val: "del_ar22", cls: "#d9ead3" },
        { lbl: "@全部割(保証抜)", type: "del_arari_avg", val: "del_ar24", cls: "#d9ead3" },
        { lbl: "@全部割(保証込)", type: "del_arari_avg", val: "del_ar25", cls: "#d9ead3" },
        { lbl: "総粗利(保証抜)", type: "del_arari_sum", val: "del_ar24", cls: "#d9ead3" },
        { lbl: "総粗利(保証込)", type: "del_arari_sum", val: "del_ar25", cls: "#d9ead3" },
        { lbl: "納車台数予算", m: "budget_del", type: "total_only", cls: "#d9ead3" },
        { lbl: "達成率", type: "total_ratio", n: "del_cnt", d: "budget_del", cls: "#d9ead3" },
        { lbl: "総粗利保証込予算", m: "budget_arari", type: "total_only_money", cls: "#d9ead3" },
        { lbl: "達成率", type: "total_ratio", n: "del_ar25", d: "budget_arari", cls: "#d9ead3" }
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

function renderCell(s, r, isT) {
    var kVal = "-", fVal = "-", tVal = "-", bg = r.cls || "#ffffff", textClass = r.redText ? "force-red" : "";
    var c = isT ? "sticky-col-total " : "";
    
    if(r.type === "total_only") {
        tVal = (s[r.m] || 0).toLocaleString();
        return "<td class='"+c+"' style='background-color:"+bg+"'><div class='cell-stack' style='align-items:center; font-weight:bold; font-size:12px; color:#444;'>" + tVal + "</div></td>";
    }
    else if(r.type === "total_only_money") {
        tVal = (s[r.m] || 0).toLocaleString();
        return "<td class='"+c+"' style='background-color:"+bg+"'><div class='cell-stack' style='align-items:center; font-weight:bold; font-size:12px; color:#444;'>¥" + tVal + "</div></td>";
    }
    else if(r.type === "total_ratio") {
        var actual = (s[r.n+"_k"] || 0) + (s[r.n+"_f"] || 0);
        var budget = s[r.d] || 0;
        tVal = budget > 0 ? Math.round((actual / budget) * 100) + "%" : "0%";
        var ratioColor = (budget > 0 && (actual / budget) >= 1) ? "color: #d32f2f;" : "color: #333;"; 
        return "<td class='"+c+"' style='background-color:"+bg+"'><div class='cell-stack' style='align-items:center; font-weight:bold; font-size:12px; " + ratioColor + "'>" + tVal + "</div></td>";
    }

    if(r.m === "empty") { return "<td class='"+c+"' style='background-color:"+bg+"'><div class='cell-stack'><div class='stack-upper' style='display:flex;'><div class='val-kei'>-</div><div class='val-fu'>-</div></div><div class='stack-lower'>-</div></div></td>"; }
    
    // 受注時の粗利ロジック
    else if(r.type && r.type.startsWith("arari")) {
        if(r.type === "arari_val") { kVal = s.j_k; fVal = s.j_f; tVal = kVal + fVal; }
        else if(r.type === "arari_avg") { var kA = s.ar_cnt_k ? Math.round(s[r.val+"_k"]/s.ar_cnt_k) : 0; var fA = s.ar_cnt_f ? Math.round(s[r.val+"_f"]/s.ar_cnt_f) : 0; var tA = (s.ar_cnt_k+s.ar_cnt_f) ? Math.round((s[r.val+"_k"]+s[r.val+"_f"])/(s.ar_cnt_k+s.ar_cnt_f)) : 0; kVal = kA.toLocaleString(); fVal = fA.toLocaleString(); tVal = tA.toLocaleString(); }
        else { kVal = Math.round(s[r.val+"_k"]).toLocaleString(); fVal = Math.round(s[r.val+"_f"]).toLocaleString(); tVal = Math.round(s[r.val+"_k"]+s[r.val+"_f"]).toLocaleString(); }
        return "<td class='"+c+"' style='background-color:"+bg+"'><div class='cell-stack'><div class='bg-sou-upper'>"+tVal+"</div><div class='bg-sou-lower'><div class='val-kei'>"+kVal+"</div><div class='val-fu'>"+fVal+"</div></div></div></td>";
    }
    
    // ★ 追加：納車時の粗利専用ロジック（分母を del_cnt で計算）
    else if(r.type && r.type.startsWith("del_arari")) {
        if(r.type === "del_arari_val") { kVal = s.del_cnt_k; fVal = s.del_cnt_f; tVal = kVal + fVal; }
        else if(r.type === "del_arari_avg") { 
            var kA = s.del_cnt_k ? Math.round(s[r.val+"_k"]/s.del_cnt_k) : 0; 
            var fA = s.del_cnt_f ? Math.round(s[r.val+"_f"]/s.del_cnt_f) : 0; 
            var tA = (s.del_cnt_k+s.del_cnt_f) ? Math.round((s[r.val+"_k"]+s[r.val+"_f"])/(s.del_cnt_k+s.del_cnt_f)) : 0; 
            kVal = kA.toLocaleString(); fVal = fA.toLocaleString(); tVal = tA.toLocaleString(); 
        }
        else { 
            kVal = Math.round(s[r.val+"_k"]).toLocaleString(); 
            fVal = Math.round(s[r.val+"_f"]).toLocaleString(); 
            tVal = Math.round(s[r.val+"_k"]+s[r.val+"_f"]).toLocaleString(); 
        }
        return "<td class='"+c+"' style='background-color:"+bg+"'><div class='cell-stack'><div class='bg-sou-upper'>"+tVal+"</div><div class='bg-sou-lower'><div class='val-kei'>"+kVal+"</div><div class='val-fu'>"+fVal+"</div></div></div></td>";
    }

    else if(r.lbl === "実績" && !r.type) {
        kVal = s.j_k; fVal = s.j_f; tVal = kVal + fVal;
        return "<td class='"+c+"' style='background-color:"+bg+"'><div class='cell-stack'><div class='stack-label-3'><div style='width:50%;border-right:1px dotted #ccc;'>軽</div><div style='width:50%;'>普</div></div><div class='stack-values-3'><div class='val-kei'>"+kVal+"</div><div class='val-fu'>"+fVal+"</div></div><div class='stack-total-3'>"+tVal+"</div></div></td>";
    }
    else {
        if(r.type === "ratio") { var nk_k = s[r.n+"_k"], dk_k = s[r.d+"_k"], nk_f = s[r.n+"_f"], dk_f = s[r.d+"_f"]; kVal = dk_k ? Math.round(nk_k/dk_k*100)+"%" : "0%"; fVal = dk_f ? Math.round(nk_f/dk_f*100)+"%" : "0%"; tVal = (dk_k+dk_f) ? Math.round((nk_k+nk_f)/(dk_k+dk_f)*100)+"%" : "0%"; }
        else if(r.type === "custom_ratio") { var nk_k = s[r.n+"_k"], dk_k = (s[r.d_sub[0]+"_k"] - s[r.d_sub[1]+"_k"]), nk_f = s[r.n+"_f"], dk_f = (s[r.d_sub[0]+"_f"] - s[r.d_sub[1]+"_f"]); kVal = dk_k > 0 ? Math.round(nk_k/dk_k*100)+"%" : "0%"; fVal = dk_f > 0 ? Math.round(nk_f/dk_f*100)+"%" : "0%"; tVal = (dk_k+dk_f) > 0 ? Math.round((nk_k+nk_f)/(dk_k+dk_f)*100)+"%" : "0%"; }
        else if(r.type === "diff") { tVal = (s[r.n+"_k"]+s[r.n+"_f"])-(s[r.d+"_k"]+s[r.d+"_f"]); }
        else if(r.type === "sum") { tVal = 0; r.items.forEach(function(i){ tVal += (s[i+"_k"]+s[i+"_f"]); }); }
        else if(r.type === "sum_ratio") { var tk_k=0, tf_f=0; r.items.forEach(function(i){ tk_k += s[i+"_k"]; tf_f += s[i+"_f"]; }); kVal = s[r.d+"_k"] ? Math.round(tk_k/s[r.d+"_k"]*100)+"%" : "0%"; fVal = s[r.d+"_f"] ? Math.round(tf_f/s[r.d+"_f"]*100)+"%" : "0%"; tVal = (s[r.d+"_k"]+s[r.d+"_f"]) ? Math.round((tk_k+tf_f)/(s[r.d+"_k"]+s[r.d+"_f"])*100)+"%" : "0%"; }
        else { kVal = s[r.m+"_k"]; fVal = s[r.m+"_f"]; tVal = kVal + fVal; }
        return "<td class='"+c+"' style='background-color:"+bg+"'><div class='cell-stack "+textClass+"'><div class='stack-upper' style='display:flex;'><div class='val-kei'>"+kVal+"</div><div class='val-fu'>"+fVal+"</div></div><div class='stack-lower'>"+tVal+"</div></div></td>";
    }
}

function showPage(id) { document.querySelectorAll('.page-content').forEach(function(p){ p.classList.remove('active'); }); document.getElementById(id).classList.add('active'); document.querySelectorAll('.tab-btn').forEach(function(b){ b.classList.remove('active'); }); var btnId = 'btn-' + id.replace('-page', ''); var btn = document.getElementById(btnId); if(btn) btn.classList.add('active'); document.getElementById('group-filter-area').style.display = (id === 'store-page') ? 'flex' : 'none'; document.getElementById('staff-filter-area').style.display = (id === 'staff-page') ? 'flex' : 'none'; }
function filterStoreByGroup() { var g = document.getElementById('group-selector').value, f = {}, t = createStats(); Object.keys(storeStatsMaster).forEach(function(st){ if (g === "all" || storeGroupMap[st] === g) { f[st] = storeStatsMaster[st]; Object.keys(t).forEach(function(k){ if(typeof t[k] === 'number') t[k] += storeStatsMaster[st][k]; }); } }); document.getElementById("store-table-container").innerHTML = buildTable(f, "店舗名", t); }
function filterStaffByStore() { var s = document.getElementById('store-selector').value, f = {}, t = createStats(); Object.keys(staffStatsMaster).forEach(function(n){ if (s === "all" || staffStoreMap[n] === s) { f[n] = staffStatsMaster[n]; Object.keys(t).forEach(function(k){ if(typeof t[k] === 'number') t[k] += staffStatsMaster[n][k]; }); } }); document.getElementById("staff-table-container").innerHTML = buildTable(f, "担当者名", t); }
