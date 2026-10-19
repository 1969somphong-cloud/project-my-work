```html
<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Maintenance Interactive Dashboard</title>

  <!-- Chart.js -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.5.0"></script>

  <!-- Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <!-- PDF -->
  <script src="https://cdn.jsdelivr.net/npm/jspdf@3.0.1/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@5.0.2/dist/jspdf.plugin.autotable.min.js"></script>

  <style>
    * {
      box-sizing: border-box;
    }

    :root {
      --glass: rgba(255,255,255,.12);
      --glass-hover: rgba(255,255,255,.18);
      --border: rgba(255,255,255,.25);
      --text: #fff;
      --muted: rgba(255,255,255,.72);
      --shadow: 0 18px 50px rgba(0,0,0,.22);
    }

    body {
      margin: 0;
      min-height: 100vh;
      color: var(--text);
      font-family:
        Inter,
        "Segoe UI",
        "Noto Sans Thai",
        sans-serif;

      background:
        radial-gradient(circle at 10% 20%, #7c3aed 0, transparent 32%),
        radial-gradient(circle at 90% 10%, #06b6d4 0, transparent 30%),
        radial-gradient(circle at 70% 90%, #ec4899 0, transparent 32%),
        linear-gradient(135deg,#111827,#1e1b4b,#0f172a);

      background-attachment: fixed;
    }

    body::before,
    body::after {
      content: "";
      position: fixed;
      width: 300px;
      height: 300px;
      border-radius: 50%;
      filter: blur(80px);
      opacity: .35;
      pointer-events: none;
      z-index: -1;
    }

    body::before {
      background: #22d3ee;
      left: -100px;
      top: 20%;
    }

    body::after {
      background: #f472b6;
      right: -100px;
      bottom: 5%;
    }

    .container {
      width: min(1400px, calc(100% - 32px));
      margin: auto;
      padding: 30px 0 50px;
    }

    /* =========================
       Glassmorphism
    ========================== */

    .glass {
      position: relative;
      overflow: hidden;

      background: var(--glass);

      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);

      border: 1px solid var(--border);
      border-radius: 22px;

      box-shadow: var(--shadow);
    }

    .glass::before {
      content: "";
      position: absolute;
      inset: 0;
      background: rgba(0,0,0,.06);
      pointer-events: none;
    }

    .glass > * {
      position: relative;
      z-index: 1;
    }

    /* =========================
       Header
    ========================== */

    .header {
      padding: 22px 26px;
      margin-bottom: 20px;

      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 20px;
    }

    .title h1 {
      margin: 0;
      font-size: clamp(24px,4vw,36px);
      font-weight: 800;
    }

    .title p {
      margin: 6px 0 0;
      color: var(--muted);
      font-size: 13px;
    }

    .header-actions {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }

    select,
    button {
      min-height: 42px;
      border-radius: 12px;
      border: 1px solid rgba(255,255,255,.25);
      background: rgba(255,255,255,.12);
      color: white;
      padding: 9px 14px;
      font: inherit;
      cursor: pointer;
    }

    select option {
      color: #111827;
    }

    button {
      font-weight: 700;
      transition: .18s ease;
    }

    button:hover {
      background: rgba(255,255,255,.22);
      transform: translateY(-1px);
    }

    button:focus-visible,
    select:focus-visible {
      outline: 2px solid white;
      outline-offset: 3px;
    }

    .btn-excel {
      background: rgba(34,197,94,.16);
    }

    .btn-pdf {
      background: rgba(239,68,68,.16);
    }

    /* =========================
       KPI
    ========================== */

    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(4,1fr);
      gap: 16px;
      margin-bottom: 20px;
    }

    .kpi {
      padding: 20px;
      min-height: 140px;
      transition: .18s ease;
    }

    .kpi:hover {
      background: var(--glass-hover);
      transform: translateY(-3px);
    }

    .kpi-label {
      color: var(--muted);
      font-size: 12px;
    }

    .kpi-value {
      margin-top: 8px;
      font-size: 32px;
      font-weight: 800;
    }

    .kpi-sub {
      margin-top: 7px;
      font-size: 11px;
      color: #bbf7d0;
    }

    .risk-high {
      color: #fca5a5;
    }

    .risk-medium {
      color: #fde68a;
    }

    .risk-low {
      color: #86efac;
    }

    /* =========================
       Charts
    ========================== */

    .chart-grid {
      display: grid;
      grid-template-columns: 1.4fr 1fr;
      gap: 20px;
      margin-bottom: 20px;
    }

    .panel {
      padding: 22px;
    }

    .panel-title {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 16px;
      gap: 12px;
    }

    .panel-title h2 {
      margin: 0;
      font-size: 17px;
    }

    .panel-title span {
      color: var(--muted);
      font-size: 11px;
    }

    .chart-container {
      height: 330px;
      position: relative;
    }

    canvas {
      max-width: 100%;
    }

    /* =========================
       Table
    ========================== */

    .table-panel {
      padding: 22px;
    }

    .table-wrapper {
      overflow-x: auto;
    }

    table {
      width: 100%;
      min-width: 760px;
      border-collapse: collapse;
    }

    th,
    td {
      padding: 13px 10px;
      text-align: left;
      border-bottom: 1px solid rgba(255,255,255,.10);
      font-size: 12px;
    }

    th {
      color: var(--muted);
      font-weight: 600;
    }

    tr:hover td {
      background: rgba(255,255,255,.05);
    }

    .badge {
      display: inline-block;
      padding: 4px 9px;
      border-radius: 999px;
      font-size: 10px;
      font-weight: 800;
    }

    .high {
      color: #fecaca;
      background: rgba(239,68,68,.18);
    }

    .medium {
      color: #fef08a;
      background: rgba(234,179,8,.18);
    }

    .low {
      color: #bbf7d0;
      background: rgba(34,197,94,.18);
    }

    /* =========================
       Footer
    ========================== */

    .footer {
      margin-top: 14px;
      color: var(--muted);
      text-align: right;
      font-size: 11px;
    }

    /* =========================
       Responsive
    ========================== */

    @media(max-width:1000px) {

      .kpi-grid {
        grid-template-columns: repeat(2,1fr);
      }

      .chart-grid {
        grid-template-columns: 1fr;
      }
    }

    @media(max-width:600px) {

      .container {
        width: calc(100% - 20px);
        padding-top: 15px;
      }

      .header {
        align-items: flex-start;
        flex-direction: column;
      }

      .header-actions {
        width: 100%;
      }

      select,
      button {
        flex: 1;
      }

      .kpi-grid {
        grid-template-columns: 1fr;
      }

      .chart-container {
        height: 280px;
      }
    }

    @media(prefers-reduced-motion:reduce) {

      *,
      *::before,
      *::after {
        animation-duration: .01ms !important;
        transition-duration: .01ms !important;
      }
    }

    @media(prefers-reduced-transparency:reduce) {

      .glass {
        background: rgba(15,23,42,.95);
        backdrop-filter: none;
        -webkit-backdrop-filter: none;
      }
    }
  </style>
</head>

<body>

<div class="container">

  <!-- =========================
       HEADER
  ========================== -->

  <header class="glass header">

    <div class="title">
      <h1>Maintenance Management Dashboard</h1>
      <p>Breakdown • Trend • Pareto • Risk Analysis</p>
    </div>

    <div class="header-actions">

      <select id="monthFilter" aria-label="Filter เดือน">

        <option value="all">
          ทุกเดือน
        </option>

        <option value="1">มกราคม</option>
        <option value="2">กุมภาพันธ์</option>
        <option value="3">มีนาคม</option>
        <option value="4">เมษายน</option>
        <option value="5">พฤษภาคม</option>
        <option value="6">มิถุนายน</option>
        <option value="7">กรกฎาคม</option>
        <option value="8">สิงหาคม</option>
        <option value="9">กันยายน</option>
        <option value="10">ตุลาคม</option>
        <option value="11">พฤศจิกายน</option>
        <option value="12">ธันวาคม</option>

      </select>

      <button id="excelBtn" class="btn-excel">
        Excel
      </button>

      <button id="pdfBtn" class="btn-pdf">
        PDF
      </button>

    </div>

  </header>


  <!-- =========================
       KPI
  ========================== -->

  <section class="kpi-grid">

    <div class="glass kpi">

      <div class="kpi-label">
        Total Breakdown
      </div>

      <div
        class="kpi-value"
        id="totalBreakdown">
        0
      </div>

      <div class="kpi-sub">
        จำนวน Breakdown
      </div>

    </div>


    <div class="glass kpi">

      <div class="kpi-label">
        Total Downtime
      </div>

      <div
        class="kpi-value"
        id="totalDowntime">
        0 h
      </div>

      <div class="kpi-sub">
        เวลาหยุดเครื่อง
      </div>

    </div>


    <div class="glass kpi">

      <div class="kpi-label">
        Average MTTR
      </div>

      <div
        class="kpi-value"
        id="averageMttr">
        0 h
      </div>

      <div class="kpi-sub">
        Downtime / Breakdown
      </div>

    </div>


    <div class="glass kpi">

      <div class="kpi-label">
        Risk Level
      </div>

      <div
        class="kpi-value"
        id="riskLevel">
        -
      </div>

      <div class="kpi-sub">
        ระดับความเสี่ยง
      </div>

    </div>

  </section>


  <!-- =========================
       CHARTS
  ========================== -->

  <section class="chart-grid">

    <!-- Trend -->

    <article class="glass panel">

      <div class="panel-title">

        <h2>
          Monthly Breakdown Trend
        </h2>

        <span>
          Trend Analysis
        </span>

      </div>

      <div class="chart-container">

        <canvas id="trendChart"></canvas>

      </div>

    </article>


    <!-- Pareto -->

    <article class="glass panel">

      <div class="panel-title">

        <h2>
          Pareto Breakdown
        </h2>

        <span>
          Machine Ranking
        </span>

      </div>

      <div class="chart-container">

        <canvas id="paretoChart"></canvas>

      </div>

    </article>

  </section>


  <!-- =========================
       TABLE
  ========================== -->

  <section class="glass table-panel">

    <div class="panel-title">

      <h2>
        Maintenance Records
      </h2>

      <span id="recordCount">
        0 Records
      </span>

    </div>

    <div class="table-wrapper">

      <table id="maintenanceTable">

        <thead>

          <tr>
            <th>เดือน</th>
            <th>เครื่องจักร</th>
            <th>ปัญหา</th>
            <th>Breakdown</th>
            <th>Downtime</th>
            <th>Risk</th>
          </tr>

        </thead>

        <tbody id="tableBody"></tbody>

      </table>

    </div>

  </section>


  <div class="footer">
    Maintenance Analytics Dashboard
  </div>

</div>


<script>

"use strict";

/* =========================================
   DATA
   =========================================

   เปลี่ยนข้อมูลส่วนนี้เป็นข้อมูลจริงของคุณได้เลย
*/

const maintenanceData = [

  {
    month: 1,
    monthName: "ม.ค.",
    machine: "M-01",
    problem: "Motor Overheat",
    breakdown: 5,
    downtime: 18.4,
    risk: "High"
  },

  {
    month: 1,
    monthName: "ม.ค.",
    machine: "M-05",
    problem: "Bearing Noise",
    breakdown: 3,
    downtime: 7.2,
    risk: "Medium"
  },

  {
    month: 2,
    monthName: "ก.พ.",
    machine: "M-01",
    problem: "Motor Overheat",
    breakdown: 4,
    downtime: 13.1,
    risk: "High"
  },

  {
    month: 2,
    monthName: "ก.พ.",
    machine: "M-08",
    problem: "Sensor Error",
    breakdown: 2,
    downtime: 2.8,
    risk: "Low"
  },

  {
    month: 3,
    monthName: "มี.ค.",
    machine: "M-05",
    problem: "Bearing Noise",
    breakdown: 6,
    downtime: 14.5,
    risk: "High"
  },

  {
    month: 3,
    monthName: "มี.ค.",
    machine: "M-12",
    problem: "Belt Damage",
    breakdown: 3,
    downtime: 6.4,
    risk: "Medium"
  },

  {
    month: 4,
    monthName: "เม.ย.",
    machine: "M-01",
    problem: "Motor Overheat",
    breakdown: 7,
    downtime: 22.2,
    risk: "High"
  },

  {
    month: 4,
    monthName: "เม.ย.",
    machine: "M-08",
    problem: "Sensor Error",
    breakdown: 2,
    downtime: 2.1,
    risk: "Low"
  },

  {
    month: 5,
    monthName: "พ.ค.",
    machine: "M-05",
    problem: "Bearing Noise",
    breakdown: 5,
    downtime: 12.7,
    risk: "High"
  },

  {
    month: 5,
    monthName: "พ.ค.",
    machine: "M-12",
    problem: "Belt Damage",
    breakdown: 4,
    downtime: 8.8,
    risk: "Medium"
  },

  {
    month: 6,
    monthName: "มิ.ย.",
    machine: "M-01",
    problem: "Motor Overheat",
    breakdown: 6,
    downtime: 19.3,
    risk: "High"
  },

  {
    month: 6,
    monthName: "มิ.ย.",
    machine: "M-08",
    problem: "Sensor Error",
    breakdown: 3,
    downtime: 3.7,
    risk: "Low"
  },

  {
    month: 7,
    monthName: "ก.ค.",
    machine: "M-05",
    problem: "Bearing Noise",
    breakdown: 4,
    downtime: 10.4,
    risk: "Medium"
  },

  {
    month: 7,
    monthName: "ก.ค.",
    machine: "M-12",
    problem: "Belt Damage",
    breakdown: 5,
    downtime: 11.1,
    risk: "High"
  },

  {
    month: 8,
    monthName: "ส.ค.",
    machine: "M-01",
    problem: "Motor Overheat",
    breakdown: 3,
    downtime: 9.5,
    risk: "Medium"
  },

  {
    month: 8,
    monthName: "ส.ค.",
    machine: "M-08",
    problem: "Sensor Error",
    breakdown: 2,
    downtime: 2.2,
    risk: "Low"
  },

  {
    month: 9,
    monthName: "ก.ย.",
    machine: "M-05",
    problem: "Bearing Noise",
    breakdown: 6,
    downtime: 15.6,
    risk: "High"
  },

  {
    month: 9,
    monthName: "ก.ย.",
    machine: "M-12",
    problem: "Belt Damage",
    breakdown: 3,
    downtime: 6.7,
    risk: "Medium"
  },

  {
    month: 10,
    monthName: "ต.ค.",
    machine: "M-01",
    problem: "Motor Overheat",
    breakdown: 4,
    downtime: 12.8,
    risk: "Medium"
  },

  {
    month: 10,
    monthName: "ต.ค.",
    machine: "M-08",
    problem: "Sensor Error",
    breakdown: 2,
    downtime: 2.5,
    risk: "Low"
  },

  {
    month: 11,
    monthName: "พ.ย.",
    machine: "M-05",
    problem: "Bearing Noise",
    breakdown: 5,
    downtime: 13.2,
    risk: "High"
  },

  {
    month: 11,
    monthName: "พ.ย.",
    machine: "M-12",
    problem: "Belt Damage",
    breakdown: 2,
    downtime: 4.3,
    risk: "Low"
  },

  {
    month: 12,
    monthName: "ธ.ค.",
    machine: "M-01",
    problem: "Motor Overheat",
    breakdown: 3,
    downtime: 8.6,
    risk: "Medium"
  },

  {
    month: 12,
    monthName: "ธ.ค.",
    machine: "M-05",
    problem: "Bearing Noise",
    breakdown: 4,
    downtime: 9.8,
    risk: "Medium"
  }

];


/* =========================================
   MONTHS
   ========================================= */

const months = [
  "ม.ค.",
  "ก.พ.",
  "มี.ค.",
  "เม.ย.",
  "พ.ค.",
  "มิ.ย.",
  "ก.ค.",
  "ส.ค.",
  "ก.ย.",
  "ต.ค.",
  "พ.ย.",
  "ธ.ค."
];


/* =========================================
   CHART INSTANCES
   ========================================= */

let trendChart = null;
let paretoChart = null;


/* =========================================
   FILTER DATA
   ========================================= */

function getFilteredData() {

  const month =
    document.getElementById("monthFilter").value;

  if (month === "all") {
    return maintenanceData;
  }

  return maintenanceData.filter(
    item => String(item.month) === month
  );
}


/* =========================================
   KPI
   ========================================= */

function updateKPI(data) {

  const totalBreakdown =
    data.reduce(
      (sum,item) => sum + item.breakdown,
      0
    );

  const totalDowntime =
    data.reduce(
      (sum,item) => sum + item.downtime,
      0
    );

  const mttr =
    totalBreakdown > 0
      ? totalDowntime / totalBreakdown
      : 0;


  /* Risk */

  const high =
    data.filter(x => x.risk === "High").length;

  const medium =
    data.filter(x => x.risk === "Medium").length;


  let risk = "LOW";
  let riskClass = "risk-low";

  if (high > 0) {

    risk = "HIGH";
    riskClass = "risk-high";

  } else if (medium > 0) {

    risk = "MEDIUM";
    riskClass = "risk-medium";

  }


  document.getElementById(
    "totalBreakdown"
  ).textContent = totalBreakdown;


  document.getElementById(
    "totalDowntime"
  ).textContent =
    totalDowntime.toFixed(1) + " h";


  document.getElementById(
    "averageMttr"
  ).textContent =
    mttr.toFixed(1) + " h";


  const riskElement =
    document.getElementById("riskLevel");

  riskElement.textContent = risk;

  riskElement.className =
    "kpi-value " + riskClass;


  document.getElementById(
    "recordCount"
  ).textContent =
    data.length + " Records";
}


/* =========================================
   TREND CHART
   ========================================= */

function createTrendChart(data) {

  const values =
    months.map((_,index) => {

      return data
        .filter(item => item.month === index + 1)
        .reduce(
          (sum,item) =>
            sum + item.breakdown,
          0
        );

    });


  if (trendChart) {
    trendChart.destroy();
  }


  const ctx =
    document
      .getElementById("trendChart")
      .getContext("2d");


  const gradient =
    ctx.createLinearGradient(
      0,
      0,
      0,
      330
    );

  gradient.addColorStop(
    0,
    "rgba(168,85,247,.45)"
  );

  gradient.addColorStop(
    1,
    "rgba(59,130,246,.02)"
  );


  trendChart =
    new Chart(ctx, {

      type: "line",

      data: {

        labels: months,

        datasets: [

          {

            label: "Breakdown",

            data: values,

            borderColor: "#c084fc",

            backgroundColor: gradient,

            borderWidth: 3,

            fill: true,

            tension: .35,

            pointRadius: 4,

            pointBackgroundColor: "#fff",

            pointBorderColor: "#a855f7",

            pointBorderWidth: 2

          }

        ]

      },

      options: {

        responsive: true,

        maintainAspectRatio: false,

        interaction: {
          mode: "index",
          intersect: false
        },

        plugins: {

          legend: {
            labels: {
              color: "#fff"
            }
          },

          tooltip: {

            callbacks: {

              label: context =>
                ` Breakdown: ${context.raw}`

            }

          }

        },

        scales: {

          x: {

            ticks: {
              color: "rgba(255,255,255,.7)"
            },

            grid: {
              color: "rgba(255,255,255,.07)"
            }

          },

          y: {

            beginAtZero: true,

            ticks: {
              color: "rgba(255,255,255,.7)"
            },

            grid: {
              color: "rgba(255,255,255,.07)"
            }

          }

        }

      }

    });

}


/* =========================================
   PARETO
   ========================================= */

function createParetoChart(data) {

  const machineMap = {};


  data.forEach(item => {

    machineMap[item.machine] =
      (machineMap[item.machine] || 0)
      + item.breakdown;

  });


  const sorted =
    Object.entries(machineMap)
      .sort((a,b) => b[1] - a[1]);


  const labels =
    sorted.map(item => item[0]);

  const values =
    sorted.map(item => item[1]);


  const total =
    values.reduce(
      (sum,value) => sum + value,
      0
    );


  let cumulative = 0;

  const cumulativePercent =
    values.map(value => {

      cumulative += value;

      return total
        ? Number(
            (cumulative / total * 100)
              .toFixed(1)
          )
        : 0;

    });


  if (paretoChart) {
    paretoChart.destroy();
  }


  const ctx =
    document
      .getElementById("paretoChart")
      .getContext("2d");


  paretoChart =
    new Chart(ctx, {

      type: "bar",

      data: {

        labels,

        datasets: [

          {

            type: "bar",

            label: "Breakdown",

            data: values,

            backgroundColor: [
              "#f87171",
              "#fbbf24",
              "#60a5fa",
              "#34d399",
              "#c084fc"
            ],

            borderRadius: 8,

            yAxisID: "y"

          },

          {

            type: "line",

            label: "Cumulative %",

            data: cumulativePercent,

            borderColor: "#fff",

            backgroundColor: "#fff",

            borderWidth: 2,

            pointRadius: 4,

            tension: .25,

            yAxisID: "percentage"

          }

        ]

      },

      options: {

        responsive: true,

        maintainAspectRatio: false,

        interaction: {
          mode: "index",
          intersect: false
        },

        plugins: {

          legend: {
            labels: {
              color: "#fff"
            }
          }

        },

        scales: {

          x: {

            ticks: {
              color: "rgba(255,255,255,.8)"
            },

            grid: {
              display: false
            }

          },

          y: {

            beginAtZero: true,

            ticks: {
              color: "rgba(255,255,255,.7)"
            },

            grid: {
              color: "rgba(255,255,255,.07)"
            }

          },

          percentage: {

            position: "right",

            min: 0,

            max: 100,

            ticks: {

              color: "#fff",

              callback: value =>
                value + "%"

            },

            grid: {
              drawOnChartArea: false
            }

          }

        }

      }

    });

}


/* =========================================
   TABLE
   ========================================= */

function updateTable(data) {

  const tbody =
    document.getElementById("tableBody");


  tbody.innerHTML = "";


  data.forEach(item => {

    const row =
      document.createElement("tr");


    let riskClass =
      item.risk.toLowerCase();


    row.innerHTML = `

      <td>
        ${item.monthName}
      </td>

      <td>
        <strong>
          ${item.machine}
        </strong>
      </td>

      <td>
        ${item.problem}
      </td>

      <td>
        ${item.breakdown}
      </td>

      <td>
        ${item.downtime.toFixed(1)} h
      </td>

      <td>

        <span class="badge ${riskClass}">
          ${item.risk}
        </span>

      </td>

    `;


    tbody.appendChild(row);

  });

}


/* =========================================
   RENDER DASHBOARD
   ========================================= */

function renderDashboard() {

  const data =
    getFilteredData();


  updateKPI(data);

  updateTable(data);

  createTrendChart(data);

  createParetoChart(data);

}


/* =========================================
   MONTH FILTER
   ========================================= */

document
  .getElementById("monthFilter")
  .addEventListener(
    "change",
    renderDashboard
  );


/* =========================================
   EXPORT EXCEL
   ========================================= */

document
  .getElementById("excelBtn")
  .addEventListener(
    "click",
    () => {

      const data =
        getFilteredData();


      const exportData =
        data.map(item => ({

          เดือน:
            item.monthName,

          เครื่องจักร:
            item.machine,

          ปัญหา:
            item.problem,

          Breakdown:
            item.breakdown,

          Downtime_Hours:
            item.downtime,

          Risk:
            item.risk

        }));


      const worksheet =
        XLSX.utils.json_to_sheet(
          exportData
        );


      const workbook =
        XLSX.utils.book_new();


      XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "Maintenance"
      );


      XLSX.writeFile(
        workbook,
        "Maintenance_Dashboard.xlsx"
      );

    }
  );


/* =========================================
   EXPORT PDF
   ========================================= */

document
  .getElementById("pdfBtn")
  .addEventListener(
    "click",
    () => {

      const data =
        getFilteredData();


      const {
        jsPDF
      } = window.jspdf;


      const pdf =
        new jsPDF({

          orientation: "landscape",

          unit: "mm",

          format: "a4"

        });


      const totalBreakdown =
        data.reduce(
          (sum,item) =>
            sum + item.breakdown,
          0
        );


      const totalDowntime =
        data.reduce(
          (sum,item) =>
            sum + item.downtime,
          0
        );


      const mttr =
        totalBreakdown
          ? totalDowntime / totalBreakdown
          : 0;


      pdf.setFontSize(18);

      pdf.text(
        "Maintenance Management Dashboard",
        14,
        16
      );


      pdf.setFontSize(10);

      pdf.text(
        `Breakdown: ${totalBreakdown} | Downtime: ${totalDowntime.toFixed(1)} h | MTTR: ${mttr.toFixed(1)} h`,
        14,
        23
      );


      pdf.autoTable({

        startY: 30,

        head: [[
          "Month",
          "Machine",
          "Problem",
          "Breakdown",
          "Downtime",
          "Risk"
        ]],

        body:
          data.map(item => [

            item.monthName,

            item.machine,

            item.problem,

            item.breakdown,

            `${item.downtime.toFixed(1)} h`,

            item.risk

          ]),

        theme: "grid",

        styles: {
          fontSize: 8
        },

        headStyles: {
          fillColor: [80,60,130]
        }

      });


      pdf.save(
        "Maintenance_Dashboard.pdf"
      );

    }
  );


/* =========================================
   INITIALIZE
   ========================================= */

renderDashboard();

</script>

</body>
</html>
```
