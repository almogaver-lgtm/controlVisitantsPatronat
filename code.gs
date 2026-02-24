/**
 * Control de Visitants per a Cellers de l'Empordà — Google Apps Script
 * Autor original: David Roca Puig · almogaver@gmail.com
 * Dashboard extension: Antigravity AI
 *
 * IMPORTANT: Substitueix TOT el codi del teu Apps Script per aquest arxiu.
 * Després: Deploy → New Deployment (o Manage Deployments → Edit → Version nova)
 *
 * Estructura del Sheet (fila 7 en endavant):
 *   Col A(0):  Data (dd/MM/yyyy)
 *   Col B(1):  Hora (HH:mm)
 *   Col C(2):  Nombre persones
 *   Col D(3):  Franja edat
 *   Col E(4):  Professional (n)
 *   Col F(5):  Amb coneixements (n)
 *   Col G(6):  Sense coneixements (n)
 *   Col H(7):  Despesa (€)
 *   Col I(8):  Catalunya (n)
 *   Col J(9):  Espanya (n)
 *   Col K(10): França (n)
 *   Col L(11): UK (n)
 *   Col M(12): Alemanya (n)
 *   Col N(13): Holanda (n)
 *   Col O(14): Nòrdics (n)
 *   Col P(15): Belgica (n)
 *   Col Q(16): Resta d'Europa (n)
 *   Col R(17): Russia (n)
 *   Col S(18): USA i Canadà (n)
 *   Col T(19): Altres país (text)
 *   Col U(20): Lliure (n)
 *   Col V(21): Allotjats al Mas (n)
 *   Col W(22): Experiències via (n)
 *   Col X(23): Recomanats (n)
 *   Col Y(24): (reservat)
 *   Col Z(25): VT (n)
 *   Col AA(26): EVT (n)
 *   Col AB(27): VTM (n)
 *   Col AC(28): MASOS (n)
 *   Col AD(29): Experiències dest (n)
 *   Col AE(30): Wine Bar (n)
 *   Col AF(31): Botiga (n)
 *   Col AG(32): Comarca
 *   Col AH(33): Població
 *   Col AI(34): Observacions
 *
 * Caselles resum:
 *   F3 = Comptador visitants avui (fórmula SUMIF)
 *   I3 = Total absolut acumulat (fórmula SUM)
 */

// ============================================================
//  doGet — AMPLIAT PER AL DASHBOARD
// ============================================================
function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var tz = ss.getSpreadsheetTimeZone();
  var now = new Date();
  var todayStr = Utilities.formatDate(now, tz, "dd/MM/yyyy");

  // KPIs principals (caselles fixes)
  var totalAvui = sheet.getRange("F3").getValue();
  var totalAbsolut = sheet.getRange("I3").getValue();

  // Llegim TOTES les dades
  var data = sheet.getDataRange().getValues();
  var row4 = data[3]; // Fila 4 (índex 3) on hi ha els sumatoris
  var masterTotals = {
    "VT":    Number(row4[25]) || 0,
    "EVT":   Number(row4[26]) || 0,
    "VTM":   Number(row4[27]) || 0,
    "MASOS": Number(row4[28]) || 0
  };

  var logs = data.slice(6); // Fila 7 en endavant

  // Definim rangs temporals
  var w1ago = new Date(now.getTime() - 7 * 86400000);
  var m1ago = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
  var q1ago = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate());

  // --- INICIALITZACIÓ ESTADÍSTIQUES PER RANG ---
  function createStatsObj() {
    return {
      hourly: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
      ages: { "18-35": 0, "36-50": 0, "51-65": 0, "+65": 0 },
      origins: {},
      business: { "Experiències": 0, "Wine Bar": 0, "Botiga": 0 },
      experienceTypes: { "VT": 0, "EVT": 0, "VTM": 0, "MASOS": 0 }
    };
  }

  var statsRange = {
    day:     createStatsObj(),
    week:    createStatsObj(),
    month:   createStatsObj(),
    quarter: createStatsObj(),
    total:   createStatsObj()
  };

  // Mantenim objecte per a KPIs (ingressos/visites) per compatibilitat
  var stats = {
    expense: { daily: 0, weekly: 0, monthly: 0, quarter: 0, total: 0 },
    counts:  { daily: 0, weekly: 0, monthly: 0, quarter: 0, total: 0 }
  };

  // Mapa de columnes de procedència dinàmic (fila 6 labels)
  var labelsRow = data[5]; 
  var paisosCols = {};
  for (var c = 8; c <= 18; c++) {
    var label = (labelsRow[c] || "").toString().trim();
    if (label) paisosCols[c] = label;
  }

  // Preparació de sèries temporals per ingressos
  var dayExpByHour = {};    // { "09": 120, "10": 80, ... }
  var weekExpByDay = {};    // { 0: sum, 1: sum, ..., 6: sum } (0=avui-6, 6=avui)
  var monthExpByWeek = [0, 0, 0, 0, 0]; // setmana 0..4
  var quarterExpByMonth = {};            // { 0: sum, 1: sum, 2: sum }
  var totalExpByYear = {};               // { "2023": sum, "2024": sum, ... }

  for (var i = 0; i < 7; i++) weekExpByDay[i] = 0;

  // Processem cada fila
  for (var r = 0; r < logs.length; r++) {
    var row = logs[r];
    var n       = Number(row[2]) || 0;
    var edat    = (row[3] || "").toString().trim();
    var despStr = (row[7] || "").toString().replace(",", ".");
    var despesa = Number(despStr) || 0;

    // --- Parsejem la DATA (pot ser Date object o string "dd/MM/yyyy") ---
    var rowDate = null;
    var rawDate = row[0];
    if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
      rowDate = rawDate;
    } else {
      var dateStr = (rawDate || "").toString().trim();
      var parts = dateStr.split("/");
      if (parts.length >= 3) {
        rowDate = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
      }
    }
    if (!rowDate) continue;  // Fila sense data vàlida, saltar

    var rowDateStr = Utilities.formatDate(rowDate, tz, "dd/MM/yyyy");

    var hora = -1;
    var rawHora = row[1];
    if (rawHora instanceof Date && !isNaN(rawHora.getTime())) {
      hora = rawHora.getHours();
    } else {
      var horaStr = (rawHora || "").toString().trim();
      if (horaStr) hora = parseInt(horaStr.split(":")[0], 10);
    }

    // Identifiquem rangs
    var isToday   = (rowDateStr === todayStr);
    var isWeekly  = (rowDate >= w1ago);
    var isMonthly = (rowDate >= m1ago);
    var isQuarter = (rowDate >= q1ago);

    // Funcio per omplir un bloc d'estadístiques per a un període
    var fillPeriod = function(s) {
      if (hora >= 0 && hora < 24) s.hourly[hora] += n;
      if (s.ages.hasOwnProperty(edat)) s.ages[edat] += n;
      
      // Orígens
      for (var colIdx in paisosCols) {
        var vVal = Number(row[colIdx]) || 0;
        if (vVal > 0) {
          var pName = paisosCols[colIdx];
          s.origins[pName] = (s.origins[pName] || 0) + vVal;
        }
      }
      // Altres països
      var altresPais = (row[19] || "").toString().trim();
      if (altresPais !== "") {
        s.origins[altresPais] = (s.origins[altresPais] || 0) + n;
      }

      // Mix de negoci
      s.business["Experiències"] += (Number(row[29]) || 0);
      s.business["Wine Bar"]     += (Number(row[30]) || 0);
      s.business["Botiga"]       += (Number(row[31]) || 0);

      // Tipus experiència
      s.experienceTypes["VT"]    += (Number(row[25]) || 0);
      s.experienceTypes["EVT"]   += (Number(row[26]) || 0);
      s.experienceTypes["VTM"]   += (Number(row[27]) || 0);
      s.experienceTypes["MASOS"] += (Number(row[28]) || 0);
    };

    // Sempre a TOTAL
    fillPeriod(statsRange.total);
    // També comptabilitzem per a KPIs compatibles
    stats.expense.total += despesa;
    stats.counts.total  += n;
    var yearStr = rowDate.getFullYear().toString();
    totalExpByYear[yearStr] = (totalExpByYear[yearStr] || 0) + despesa;

    if (isQuarter) {
      fillPeriod(statsRange.quarter);
      stats.expense.quarter += despesa;
      stats.counts.quarter  += n;
      var mKey = rowDate.getMonth();
      quarterExpByMonth[mKey] = (quarterExpByMonth[mKey] || 0) + despesa;
    }
    if (isMonthly) {
      fillPeriod(statsRange.month);
      stats.expense.monthly += despesa;
      stats.counts.monthly  += n;
      var dM = rowDate.getDate();
      var weekIdx = Math.min(Math.floor((dM - 1) / 7), 4);
      monthExpByWeek[weekIdx] += despesa;
    }
    if (isWeekly) {
      fillPeriod(statsRange.week);
      stats.expense.weekly += despesa;
      stats.counts.weekly  += n;
      var diffDays = Math.floor((now.getTime() - rowDate.getTime()) / 86400000);
      if (diffDays >= 0 && diffDays < 7) weekExpByDay[6 - diffDays] += despesa;
    }
    if (isToday) {
      fillPeriod(statsRange.day);
      stats.expense.daily += despesa;
      stats.counts.daily  += n;
      if (hora >= 0) {
        var hKey = (hora < 10 ? "0" : "") + hora;
        dayExpByHour[hKey] = (dayExpByHour[hKey] || 0) + despesa;
      }
    }
  }

  // --- Calculem mitjanes ---
  var expenseMean = {
    daily:   stats.counts.daily   > 0 ? Math.round((stats.expense.daily   / stats.counts.daily)   * 100) / 100 : 0,
    weekly:  stats.counts.weekly  > 0 ? Math.round((stats.expense.weekly  / stats.counts.weekly)  * 100) / 100 : 0,
    monthly: stats.counts.monthly > 0 ? Math.round((stats.expense.monthly / stats.counts.monthly) * 100) / 100 : 0,
    quarter: stats.counts.quarter > 0 ? Math.round((stats.expense.quarter / stats.counts.quarter) * 100) / 100 : 0,
    total:   stats.counts.total   > 0 ? Math.round((stats.expense.total   / stats.counts.total)   * 100) / 100 : 0
  };

  // --- Construïm sèries temporals per a les gràfiques ---

  // Sèrie DIÀRIA: hores amb despesa (8h-21h)
  var dayLabels = [];
  var dayValues = [];
  for (var hh = 8; hh <= 21; hh++) {
    var hhStr = (hh < 10 ? "0" : "") + hh;
    dayLabels.push(hhStr + "h");
    dayValues.push(dayExpByHour[hhStr] || 0);
  }

  // Sèrie SETMANAL: 7 dies (Dl..Dg)
  var weekDayNames = ["Dg", "Dl", "Dt", "Dc", "Dj", "Dv", "Ds"];
  var weekLabels = [];
  for (var d = 6; d >= 0; d--) {
    var pastDate = new Date(now.getTime() - d * 86400000);
    weekLabels.push(weekDayNames[pastDate.getDay()]);
  }
  var weekValues = [];
  for (var wi = 0; wi < 7; wi++) {
    weekValues.push(weekExpByDay[wi] || 0);
  }

  // Sèrie MENSUAL: setmanes S1..S5
  var monthLabels = ["S1", "S2", "S3", "S4", "S5"];
  var monthValues = monthExpByWeek;

  // Sèrie TRIMESTRAL: noms dels mesos
  var monthNames = ["Gen", "Feb", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Des"];
  var quarterLabels = [];
  var quarterValues = [];
  for (var mi = 2; mi >= 0; mi--) {
    var mDate = new Date(now.getFullYear(), now.getMonth() - mi, 1);
    var mIdx = mDate.getMonth();
    quarterLabels.push(monthNames[mIdx]);
    quarterValues.push(quarterExpByMonth[mIdx] || 0);
  }

  // Sèrie TOTAL: per anys
  var yearLabels = Object.keys(totalExpByYear).sort();
  var yearValues = yearLabels.map(function(y) { return totalExpByYear[y]; });

  var expenseSeries = {
    day:     { labels: dayLabels,     values: dayValues },
    week:    { labels: weekLabels,    values: weekValues },
    month:   { labels: monthLabels,   values: monthValues },
    quarter: { labels: quarterLabels, values: quarterValues },
    total:   { labels: yearLabels,    values: yearValues }
  };

  // --- Resposta final ---
  var result = {
    status: "ok",
    dailyCount: totalAvui,
    absoluteTotal: totalAbsolut,
    stats: {
      expense:       stats.expense,
      expenseMean:   expenseMean,
      counts:        stats.counts,
      masterTotals:    masterTotals,
      expenseSeries:   expenseSeries,
      // Nou camp amb tot desglossat per rang
      statsRange:    statsRange
    }
  };

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
//  doPost — SENSE CANVIS (funciona perfectament)
// ============================================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return crearRespostaError("El servidor està ocupat.");
  }

  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var tz = ss.getSpreadsheetTimeZone();
    var d = new Date();
    var dataFormatada = Utilities.formatDate(d, tz, "dd/MM/yyyy");
    var horaFormatada = Utilities.formatDate(d, tz, "HH:mm");

    var n = Number(data.numPersones);
    var row = Array(35).fill("");

    row[0] = dataFormatada;
    row[1] = horaFormatada;
    row[2] = n;
    row[3] = data.edat;

    if (data.tipusClient === "Professional") row[4] = n;
    else if (data.tipusClient === "Amb coneixements") row[5] = n;
    else if (data.tipusClient === "Sense coneixements") row[6] = n;

    row[7] = data.despesa;

    var paisosMap = {
      8: "Catalunya", 9: "Espanya", 10: "França", 11: "UK", 12: "Alemanya",
      13: "Holanda", 14: "Nòrdics", 15: "Belgica", 16: "Resta d'Europa",
      17: "Russia", 18: "USA i Canadà"
    };

    var trobat = false;
    for (var colIndex in paisosMap) {
      if (data.procedencia === paisosMap[colIndex]) {
        row[colIndex] = n;
        trobat = true;
        break;
      }
    }

    if (!trobat || data.procedencia === "Altres") {
      row[19] = data.altresPais || "No especificat";
    }

    if (data.comVenen === "Allotjats al Mas") {
      row[21] = n;
    } else {
      if (data.comVenen === "Lliure") row[20] = n;
      if (data.comVenen === "Recomanats") row[23] = n;
      if (data.destinacio === "Experiències") row[22] = n;
    }

    if (data.tipusVisita === "VT") row[25] = n;
    else if (data.tipusVisita === "EVT") row[26] = n;
    else if (data.tipusVisita === "VTM") row[27] = n;
    else if (data.tipusVisita === "MASOS") row[28] = n;

    if (data.destinacio === "Experiències") row[29] = n;
    else if (data.destinacio === "Wine Bar") row[30] = n;
    else if (data.destinacio === "Botiga") row[31] = n;

    row[32] = data.comarca;
    row[33] = data.poblacio;
    row[34] = data.observacions;

    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow < 6 ? 7 : lastRow + 1, 1, 1, 35).setValues([row]);

    SpreadsheetApp.flush();
    var nouTotal = sheet.getRange("F3").getValue();

    return ContentService.createTextOutput(JSON.stringify({ 
      status: "ok",
      dailyCount: nouTotal
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return crearRespostaError(e.toString());
  } finally {
    lock.releaseLock();
  }
}

function crearRespostaError(msg) {
  return ContentService.createTextOutput(JSON.stringify({status: "error", message: msg})).setMimeType(ContentService.MimeType.JSON);
}
