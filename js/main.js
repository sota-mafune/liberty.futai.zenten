var allData = [];
var loggedInStoreName = ""; 
var targetMonth = "";

ZOHO.embeddedApp.init().then(async function() {
    try {
        const userRes = await ZOHO.CRM.CONFIG.getCurrentUser();
        if(userRes.users && userRes.users.length > 0) {
            loggedInStoreName = userRes.users[0].full_name; 
        }
    } catch(e) { console.error("ユーザー取得エラー", e); }

    initMonthSelector();
    targetMonth = document.getElementById('month-selector').value;
    fetchAllRecords(1);
});

function initMonthSelector() { 
    var sel = document.getElementById('month-selector'); 
    var now = new Date(); 
    var curMonth = now.getFullYear() + "-" + ("0" + (now.getMonth() + 1)).slice(-2); 
    for(var y=2025; y<=2026; y++) { 
        for(var m=1; m<=12; m++) { 
            var v = y + "-" + (m < 10 ? "0" + m : m); 
            var opt = document.createElement('option'); 
            opt.value = v; 
            opt.text = y + "年 " + m + "月"; 
            if(v === curMonth) opt.selected = true; 
            sel.add(opt); 
        } 
    } 
}

function changeMonth() {
    targetMonth = document.getElementById('month-selector').value;
    allData = [];
    document.getElementById('loading').style.display = 'block';
    fetchAllRecords(1);
}

function fetchAllRecords(page) {
    ZOHO.CRM.API.getAllRecords({ Entity: "Services", page: page, per_page: 200 }).then(function(res) {
        if (res.data) {
            allData = allData.concat(res.data);
            if (res.info && res.info.more_records) {
                fetchAllRecords(page + 1); 
            } else {
                finishFetch();
            }
        } else {
            finishFetch();
        }
    }).catch(function(e){
        console.error("データ取得エラー", e);
        finishFetch();
    });
}

function finishFetch() {
    updateStoreDropdown();
    renderAll();
}

function updateStoreDropdown() {
    let storeSet = new Set();
    allData.forEach(r => { 
        let st = r.ServiceStore || "未所属";
        let stName = (st && typeof st === 'object') ? (st.name || "未所属") : st;
        if (stName !== "未所属") storeSet.add(stName); 
    });

    if (loggedInStoreName && !storeSet.has(loggedInStoreName)) {
        storeSet.add(loggedInStoreName);
    }

    let sel = document.getElementById('store-selector');
    sel.innerHTML = '';
    
    let sortedStores = Array.from(storeSet).sort();
    sortedStores.forEach(st => {
        var o = document.createElement('option'); o.value = o.text = st; sel.add(o);
    });

    if (loggedInStoreName) {
        for (let i = 0; i < sel.options.length; i++) {
            if (sel.options[i].value === loggedInStoreName) {
                sel.selectedIndex = i;
                break;
            }
        }
    }
}

function createStats() { return { j_k:0, j_f:0, v_n_k:0, v_n_f:0, sho_k:0, sho_f:0, ab_k:0, ab_f:0, jk_k:0, jk_f:0, rv_k:0, rv_f:0, rj_k:0, rj_f:0, tot_v_k:0, tot_v_f:0, n_k:0, n_f:0, m_k:0, m_f:0, c_k:0, c_f:0, o2_k:0, o2_f:0, o3_k:0, o3_f:0, pk_k:0, pk_f:0, ct_k:0, ct_f:0, up_k:0, up_f:0, tp_k:0, tp_f:0, ic_k:0, ic_f:0, rst_k:0, rst_f:0, ni_k:0, ni_f:0, nu_k:0, nu_f:0, hp_k:0, hp_f:0, fl_k:0, fl_f:0, aq_k:0, aq_f:0, tr_k:0, tr_f:0, ln_k:0, ln_f:0, l84_k:0, l84_f:0, r69_k:0, r69_f:0, r59_k:0, r59_f:0, r49_k:0, r49_f:0, r39_k:0, r39_f:0, r29_k:0, r29_f:0, low_k:0, low_f:0, zn_k:0, zn_f:0, ar21_k:0, ar21_f:0, ar22_k:0, ar22_f:0, ar23_k:0, ar23_f:0, ar24_k:0, ar24_f:0, ar25_k:0, ar25_f:0, ar_cnt_k:0, ar_cnt_f:0, g_sls:0 }; }

function aggregate(s, rec) {
    var c = (rec.SyaryouCategory === "軽" || rec.SyaryouCategory === "軽自動車") ? "k" : "f";
    var vD = (rec.VisitedDateTime || "").split('T')[0], cD = (rec.ClosingDay || ""), isCancel = (rec.cancel === true || rec.cancel === "true");

    if (vD && vD.indexOf(targetMonth) !== -1) { 
        s["tot_v_"+c]++; 
        if (rec.FOrR === "初回") s["v_n_"+c]++; 
        if (rec.Seated === "〇") s["sho_"+c]++; 
        if (rec.Option1 === "〇") s["ab_"+c]++; 
        if (rec.FOrR === "初回" && cD === vD && !isCancel) s["jk_"+c]++; 
        if (rec.FOrR === "再来") s["rv_"+c]++; 
    }
    if (cD && cD.indexOf(targetMonth) !== -1 && !isCancel) { 
        s["j_"+c]++; 
        if (rec.FOrR === "再来") s["rj_"+c]++; 
        if (rec.HanbaiCategory === "新") s["n_"+c]++; 
        if (rec.HanbaiCategory === "未") s["m_"+c]++; 
        if (rec.HanbaiCategory === "中") s["c_"+c]++; 
        if (rec.Option2 === "〇") s["o2_"+c]++; 
        if (rec.Option3 === "〇") s["o3_"+c]++; 
        if (["新車ﾊﾟｯｸ","未使用ﾊﾟｯｸ","中ｴｺﾊﾟｯｸ","中ｽﾀﾊﾟｯｸ","中ｱﾌﾟﾊﾟｯｸ"].indexOf(rec.Option15) !== -1) s["pk_"+c]++; 
        if (rec.Option4 === "〇") s["ct_"+c]++; 
        if (rec.Option5 === "〇" || rec.Option21 === "〇") s["up_"+c]++;
        if (rec.Option16 === "〇") s["tp_"+c]++; 
        if (rec.Option7 === "〇") s["ic_"+c]++; 
        if (rec.Option8 === "〇") s["rst_"+c]++; 
        if (rec.Option9 === "〇") s["ni_"+c]++; 
        if (rec.Option10 === "〇") s["nu_"+c]++; 
        if (rec.Option6 === "〇") s["hp_"+c]++; 
        if (rec.BackCamera === "〇") s["fl_"+c]++; 
        if (rec.Option17 === "〇") s["aq_"+c]++; 
        if (["✖","✕","×","","null","-","無","なし"].indexOf(rec.TradeinCar || "") === -1) s["tr_"+c]++; 
        if ((rec.PaymentCategory || "").indexOf("ローン") !== -1) { 
            s["ln_"+c]++; 
            var r = parseFloat(rec.arari17)||0, ct = parseInt(rec.arari16)||0; 
            if (ct >= 84) s["l84_"+c]++; 
            if (r >= 6 && r < 7) s["r69_"+c]++; else if (r >= 5 && r < 6) s["r59_"+c]++; else if (r >= 4 && r < 5) s["r49_"+c]++; else if (r >= 3 && r < 4) s["r39_"+c]++; else if (r >= 2.9 && r < 3) s["r29_"+c]++; else if (r > 0 && r < 2.9) s["low_"+c]++; 
        } 
        if (rec.Option14 === "〇") s["zn_"+c]++; 
        s["ar21_"+c] += (parseFloat(rec.arari21)||0); s["ar22_"+c] += (parseFloat(rec.arari22)||0); s["ar23_"+c] += (parseFloat(rec.arari23)||0); s["ar24_"+c] += (parseFloat(rec.arari24)||0); s["ar25_"+c] += (parseFloat(rec.arari25)||0); s["ar_cnt_"+c]++; 
    }
}

function renderAll() {
    document.getElementById('loading').style.display = 'none';
    let selectedStore = document.getElementById('store-selector').value;
    if(!selectedStore) return;

    let staffStatsMaster = {};
    let staffStoreMap = {};
    let totalStats = createStats();
    let dailyStats = {};

    let p = targetMonth.split("-").map(Number);
    let lastDay = new Date(p[0], p[1], 0).getDate();
    for(let d=1; d<=lastDay; d++) {
        let dtStr = targetMonth + "-" + ("0"+d).slice(-2);
        dailyStats[dtStr] = createStats();
    }

    allData.forEach(r => {
        let st = r.ServiceStore || "未所属";
        let stName = (st && typeof st === 'object') ? (st.name || "未所属") : st;
        let pr = "未設定";
        if (r.ServicePerson) {
            if (typeof r.ServicePerson === 'string') pr = r.ServicePerson;
            else if (typeof r.ServicePerson === 'object') pr = r.ServicePerson.name || r.ServicePerson.full_name || "未設定";
        }
        if (!staffStatsMaster[pr]) staffStatsMaster[pr] = createStats();
        if (stName !== "未所属") staffStoreMap[pr] = stName; 
    });

    let filteredStaffStats = {};
    Object.keys(staffStatsMaster).forEach(pr => {
        if (staffStoreMap[pr] === selectedStore) {
            filteredStaffStats[pr] = createStats();
        }
    });

    allData.forEach(r => {
        let st = r.ServiceStore || "未所属";
        let stName = (st && typeof st === 'object') ? (st.name || "未所属") : st;

        if (stName !== selectedStore) return;

        let pr = "未設定";
        if (r.ServicePerson) {
            if (typeof r.ServicePerson === 'string') pr = r.ServicePerson;
            else if (typeof r.ServicePerson === 'object') pr = r.ServicePerson.name || r.ServicePerson.full_name || "未設定";
        }

        if (!filteredStaffStats[pr]) filteredStaffStats[pr] = createStats();

        aggregate(filteredStaffStats[pr], r);
        aggregate(totalStats, r);

        var c = (r.SyaryouCategory === "軽" || r.SyaryouCategory === "軽自動車") ? "k" : "f";
        var vD = (r.VisitedDateTime || "").split('T')[0], cD = (r.ClosingDay || ""), isCancel = (r.cancel === true || r.cancel === "true");

        if (vD && vD.indexOf(targetMonth) !== -1) { 
            if(dailyStats[vD]){ let ds = dailyStats[vD]; ds["tot_v_"+c]++; if (r.FOrR === "初回") ds["v_n_"+c]++; if (r.Seated === "〇") ds["sho_"+c]++; if (r.FOrR === "再来") ds["rv_"+c]++; }
        }
        if (cD && cD.indexOf(targetMonth) !== -1 && !isCancel) { 
            if(dailyStats[cD]) dailyStats[cD]["j_"+c]++; 
        }
    });

    document.getElementById("left-table-container").innerHTML = buildTable(filteredStaffStats, "担当者名", totalStats);
    document.getElementById("right-table-container").innerHTML = buildDailyTable(dailyStats, totalStats);
}

function buildTable(sum, title, totalS) {
    var keys = Object.keys(sum).sort();
    
    // 【修正1】一番上の見出し（店舗名・合計）がスクロールの下に潜り込まないように z-index を高くする！
    var h = "<table><thead><tr><th class='sticky-col-item shop-header' style='z-index: 200;'>" + title + "</th><th class='sticky-col-total shop-header' style='z-index: 190;'>合計</th>";
    for(var i=0; i<keys.length; i++) h += "<th class='shop-header'>" + keys[i] + "</th>";
    h += "</tr></thead><tbody>";

    const rowDef = [
        { sec: "予算・目標" }, { lbl: "予算", m: "empty", cls: "#ffffff" }, { lbl: "目標", m: "empty", cls: "#ffffff" }, { lbl: "昨年実績", m: "empty", cls: "#ffffff" }, { lbl: "現時点予算", m: "empty", cls: "#ffffff" },
        { sec: "基本実績" }, { lbl: "実績", m: "j", cls: "#ffe599" }, { lbl: "達成率", type: "ratio", n: "j", d: "g_sls", cls: "#ffffff" }, { lbl: "昨年実績(当日)", m: "empty", cls: "#d9d2e9" }, { lbl: "昨年対比", m: "empty", cls: "#d9d2e9" },
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
        { sec: "受注時想定" }, { lbl: "受注台数", type: "arari_val", val: "j", cls: "#ead1dc" }, { lbl: "@車両粗利", type: "arari_avg", val: "ar21", cls: "#ead1dc" }, { lbl: "@ローンBK", type: "arari_avg", val: "ar23", cls: "#ead1dc" }, { lbl: "@下取粗利", type: "arari_avg", val: "ar22", cls: "#ead1dc" }, { lbl: "@全部割(保証抜き)", type: "arari_avg", val: "ar24", cls: "#ead1dc" }, { lbl: "@全部割(保証込み)", type: "arari_avg", val: "ar25", cls: "#ead1dc" }, { lbl: "総粗利", type: "arari_sum", val: "ar25", cls: "#ead1dc" }
    ];

    for(var j=0; j<rowDef.length; j++){ 
        var r = rowDef[j]; 
        if(r.sec) {
            // 【修正2】帯を3分割して、固定列とスクロール列を完全に分離する！
            h += "<tr>";
            h += "<td class='sticky-col-item section-row' style='z-index: 180; border-right: none;'>" + r.sec + "</td>";
            h += "<td class='sticky-col-total section-row' style='z-index: 180; border-left: none;'></td>";
            if (keys.length > 0) {
                h += "<td colspan='" + keys.length + "' class='section-row' style='position: static; z-index: 1; border-left: none;'></td>";
            }
            h += "</tr>";
        }
        else {
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
    if(r.m === "empty") { return "<td class='"+c+"' style='background-color:"+bg+"'><div class='cell-stack'><div class='stack-upper' style='display:flex;'><div class='val-kei'>-</div><div class='val-fu'>-</div></div><div class='stack-lower'>-</div></div></td>"; }
    else if(r.type && r.type.startsWith("arari")) {
        if(r.type === "arari_val") { kVal = s.j_k; fVal = s.j_f; tVal = kVal + fVal; }
        else if(r.type === "arari_avg") { var kA = s.ar_cnt_k ? Math.round(s[r.val+"_k"]/s.ar_cnt_k) : 0; var fA = s.ar_cnt_f ? Math.round(s[r.val+"_f"]/s.ar_cnt_f) : 0; var tA = (s.ar_cnt_k+s.ar_cnt_f) ? Math.round((s[r.val+"_k"]+s[r.val+"_f"])/(s.ar_cnt_k+s.ar_cnt_f)) : 0; kVal = kA.toLocaleString(); fVal = fA.toLocaleString(); tVal = tA.toLocaleString(); }
        else { kVal = Math.round(s[r.val+"_k"]).toLocaleString(); fVal = Math.round(s[r.val+"_f"]).toLocaleString(); tVal = Math.round(s[r.val+"_k"]+s[r.val+"_f"]).toLocaleString(); }
        return "<td class='"+c+"'><div class='cell-stack'><div class='bg-sou-upper'>"+tVal+"</div><div class='bg-sou-lower'><div class='val-kei'>"+kVal+"</div><div class='val-fu'>"+fVal+"</div></div></div></td>";
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

function buildDailyTable(dStats, tStats) {
    var days = Object.keys(dStats).sort(), weekDays = ["日", "月", "火", "水", "木", "金", "土"];
    var h = "<table class='daily-table'><thead><tr><th rowspan='2' style='min-width:30px; left:0; z-index:110;'>日</th><th rowspan='2' style='min-width:30px; left:30px; z-index:110;'>曜</th><th rowspan='2'>新規<br>来場</th><th rowspan='2'>再<br>来場</th><th rowspan='2'>総<br>来場</th><th rowspan='2' class='bg-yellow'>予算</th><th rowspan='2' class='bg-yellow'>着地<br>予想</th><th rowspan='2'>実績</th><th rowspan='2'>予算<br>進捗</th><th rowspan='2'>着地<br>予想<br>進捗</th><th rowspan='2'>商談<br>数</th><th rowspan='2'>商談<br>率</th><th rowspan='2'>成約<br>率</th><th rowspan='2' style='background:#ccc;'>昨年<br>新規</th><th rowspan='2' style='background:#ccc;'>昨年<br>総</th><th rowspan='2' style='background:#ccc;'>昨年<br>成約</th><th colspan='3' class='th-group-kei'>軽自動車</th><th colspan='3' class='th-group-fu'>普通車</th></tr><tr><th class='th-group-kei'>新規<br>来場</th><th class='th-group-kei'>再<br>来場</th><th class='th-group-kei'>実績</th><th class='th-group-fu'>新規<br>来場</th><th class='th-group-fu'>再<br>来場</th><th class='th-group-fu'>実績</th></tr></thead><tbody>";
    days.forEach(d => {
        var s = dStats[d], dateObj = new Date(d), dayNum = dateObj.getDate(), weekIdx = dateObj.getDay(), weekStr = weekDays[weekIdx], dayClass = (weekIdx === 0) ? "day-sun" : (weekIdx === 6) ? "day-sat" : "";
        var vn = s.v_n_k + s.v_n_f, rv = s.rv_k + s.rv_f, tot = s.tot_v_k + s.tot_v_f, j = s.j_k + s.j_f, sho = s.sho_k + s.sho_f;
        var sho_r = vn > 0 ? Math.round(sho / vn * 100) + "%" : "-", j_r = vn > 0 ? Math.round(j / vn * 100) + "%" : "-";
        h += "<tr><td class='center' style='position:sticky; left:0; background:#fff; z-index:90;'>"+dayNum+"</td><td class='center "+dayClass+"' style='position:sticky; left:30px; background:#fff; z-index:90;'>"+weekStr+"</td><td>"+(vn||"")+"</td><td>"+(rv||"")+"</td><td>"+(tot||"")+"</td><td class='bg-yellow'>-</td><td class='bg-yellow'>-</td><td>"+(j||"")+"</td><td>-</td><td>-</td><td>"+(sho||"")+"</td><td>"+sho_r+"</td><td>"+j_r+"</td><td style='background:#f2f2f2;'>-</td><td style='background:#f2f2f2;'>-</td><td style='background:#f2f2f2;'>-</td><td>"+(s.v_n_k||"")+"</td><td>"+(s.rv_k||"")+"</td><td>"+(s.j_k||"")+"</td><td>"+(s.v_n_f||"")+"</td><td>"+(s.rv_f||"")+"</td><td>"+(s.j_f||"")+"</td></tr>";
    });
    var tvn = tStats.v_n_k + tStats.v_n_f, trv = tStats.rv_k + tStats.rv_f, ttot = tStats.tot_v_k + tStats.tot_v_f, tj = tStats.j_k + tStats.j_f, tsho = tStats.sho_k + tStats.sho_f;
    var tsho_r = tvn > 0 ? Math.round(tsho / tvn * 100) + "%" : "-", tj_r = tvn > 0 ? Math.round(tj / tvn * 100) + "%" : "-";
    h += "<tr class='daily-total-row'><td colspan='2' class='center' style='position:sticky; left:0; z-index:90;'>合計</td><td>"+tvn+"</td><td>"+trv+"</td><td>"+ttot+"</td><td class='bg-yellow'>-</td><td class='bg-yellow'>-</td><td>"+tj+"</td><td>-</td><td>-</td><td>"+tsho+"</td><td>"+tsho_r+"</td><td>"+tj_r+"</td><td>-</td><td>-</td><td>-</td><td>"+tStats.v_n_k+"</td><td>"+tStats.rv_k+"</td><td>"+tStats.j_k+"</td><td>"+tStats.v_n_f+"</td><td>"+tStats.rv_f+"</td><td>"+tStats.j_f+"</td></tr></tbody></table>";
    return h;
}
