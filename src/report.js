import Papa from "papaparse";
import Chart from "chart.js/auto";
import ChartDataLabels from "chartjs-plugin-datalabels";
import rawCSV from './dummy_emails.csv';
import jsPDF from "jspdf";
import './report.css';

Chart.register(ChartDataLabels);

let emailData = [];
let chartInstance = null;

document.addEventListener("DOMContentLoaded", () => {
  const currentYear = new Date().getFullYear();
  const yearSelect = document.getElementById("year-select");
  const monthYearSelect = document.getElementById("month-year-select");

  for (let y = currentYear; y >= currentYear - 5; y--) {
    const option = new Option(y, y);
    yearSelect.add(option.cloneNode(true));
    monthYearSelect.add(option);
  }

  document.getElementById("view-mode").addEventListener("change", (e) => {
    const mode = e.target.value;
    document.getElementById("custom-date-range").style.display = mode === "custom" ? "flex" : "none";
    document.getElementById("month-picker").style.display = mode === "month" ? "flex" : "none";
    document.getElementById("year-picker").style.display = mode === "year" ? "flex" : "none";
  });

  document.getElementById("generate-btn").onclick = () => {
    const visual = document.getElementById("visual-select").value;
    const groupByConversationId = document.getElementById("group-select").value === "yes";
    const viewMode = document.getElementById("view-mode").value;
    let startDate = "", endDate = "";

    if (viewMode === "month") {
      const month = parseInt(document.getElementById("month-select").value);
      const year = parseInt(document.getElementById("month-year-select").value);
      startDate = `${year}-${String(month + 1).padStart(2, '0')}-01`;
      const lastDay = new Date(year, month + 1, 0).getDate();
      endDate = `${year}-${String(month + 1).padStart(2, '0')}-${lastDay}`;
    } else if (viewMode === "year") {
      const year = parseInt(document.getElementById("year-select").value);
      startDate = `${year}-01-01`;
      endDate = `${year}-12-31`;
    } else {
      startDate = document.getElementById("start-date").value;
      endDate = document.getElementById("end-date").value;
    }

    updateChart(visual, groupByConversationId, startDate, endDate, viewMode);
  };

  Papa.parse(rawCSV, {
    header: true,
    skipEmptyLines: true,
    complete: (result) => {
      emailData = result.data;
      document.getElementById("generate-btn").click();
    }
  });
});

function updateChart(type, groupByConvo, startDate, endDate, viewMode) {
  const filteredData = filterData(emailData, groupByConvo, startDate, endDate);
  if (chartInstance) chartInstance.destroy();

  switch (type) {
    case "daily":
      if (viewMode === "month") renderDailyForMonth(filteredData, startDate);
      else if (viewMode === "year") renderMonthlyForYear(filteredData, startDate);
      else renderEmailsPerDay(filteredData);
      break;
    case "category": renderEmailsByCategory(filteredData); break;
    case "assignee": renderEmailsByAssignee(filteredData); break;
    case "attachments": renderEmailsByAttachment(filteredData); break;
  }
}

function filterData(data, groupByConvo, start, end) {
  const seenConvos = new Set();
  return data.filter(email => {
    const rawDate = (email.date || "").trim();
    const emailDatePart = rawDate.split(" ")[0];
    if (!/^\d{4}-\d{2}-\d{2}$/.test(emailDatePart)) return false;
    if (start && emailDatePart < start) return false;
    if (end && emailDatePart > end) return false;
    if (groupByConvo && seenConvos.has(email.conversationId)) return false;
    seenConvos.add(email.conversationId);
    return true;
  });
}

function renderDailyForMonth(data, startDate) {
  const [year, month] = startDate.split("-").map(Number);
  const daysInMonth = new Date(year, month, 0).getDate();
  const countMap = {};
  for (let d = 1; d <= daysInMonth; d++) {
    const key = `${year}-${String(month).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    countMap[key] = 0;
  }
  data.forEach(email => {
    const raw = email.date?.split(" ")[0];
    if (raw?.startsWith(`${year}-${String(month).padStart(2, "0")}`)) {
      countMap[raw] = (countMap[raw] || 0) + 1;
    }
  });
  const date = new Date(year, month - 1);
  renderBarChart(`Daily Emails - ${date.toLocaleString("default", { month: "long" })} ${year}`, countMap);
}

function renderMonthlyForYear(data, startDate) {
  const [year] = startDate.split("-").map(Number);
  const countMap = {};
  for (let m = 0; m < 12; m++) {
    const label = new Date(year, m).toLocaleString("default", { month: "short" });
    countMap[label] = 0;
  }
  data.forEach(email => {
    const raw = email.date?.split(" ")[0];
    if (!raw?.startsWith(`${year}-`)) return;
    const monthIndex = parseInt(raw.split("-")[1], 10) - 1;
    const label = new Date(year, monthIndex).toLocaleString("default", { month: "short" });
    countMap[label] = (countMap[label] || 0) + 1;
  });
  renderBarChart(`Monthly Emails - ${year}`, countMap);
}

function renderEmailsPerDay(data) {
  const countMap = {};
  data.forEach(email => {
    const dateStr = (email.date || "").trim().split(" ")[0];
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return;
    countMap[dateStr] = (countMap[dateStr] || 0) + 1;
  });
  renderBarChart("Emails Received Each Day", countMap);
}

function renderEmailsByCategory(data) {
  const map = {};
  data.forEach(email => {
    const cat = email.category || "Uncategorized";
    map[cat] = (map[cat] || 0) + 1;
  });
  renderPieChart("Emails by Category", map);
}

function renderEmailsByAssignee(data) {
  const map = {};
  data.forEach(email => {
    const assignee = email.assignedTo || "Unassigned";
    map[assignee] = (map[assignee] || 0) + 1;
  });
  const orderedMap = {};
  Object.keys(map).filter(k => k !== "Unassigned").sort().forEach(k => orderedMap[k] = map[k]);
  if (map["Unassigned"]) orderedMap["Unassigned"] = map["Unassigned"];
  renderBarChart("Emails by Assignee", orderedMap);
}

function renderEmailsByAttachment(data) {
  let withAttachments = 0, withoutAttachments = 0;
  data.forEach(email => {
    if (email.hasAttachments?.trim().toLowerCase() === "true") withAttachments++;
    else withoutAttachments++;
  });
  renderPieChart("Emails With vs Without Attachments", {
    "With Attachments": withAttachments,
    "Without Attachments": withoutAttachments
  });
}

function renderBarChart(title, map) {
  // Sort the entries by date (or label) ascending
  const entries = Object.entries(map).sort(
    ([a], [b]) => new Date(a) - new Date(b)
  );

  const labels = entries.map(([k]) => k);
  const values = entries.map(([_, v]) => v);

  if (labels.length === 0) {
    alert("No data available for this chart.");
    return;
  }

  const canvas = document.getElementById("main-chart");
  chartInstance = new Chart(canvas, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: title,
        data: values,
        backgroundColor: "#0078d4"
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: { display: true, text: title },
        datalabels: { display: false }
      }
    },
    plugins: [ChartDataLabels]
  });
}


function renderPieChart(title, map) {
  const labels = Object.keys(map);
  const values = Object.values(map);
  chartInstance = new Chart(document.getElementById("main-chart"), {
    type: "pie",
    data: {
      labels,
      datasets: [{ data: values, backgroundColor: ["#4caf50", "#ff9800", "#03a9f4", "#f44336", "#9c27b0"] }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: { display: true, text: title },
        datalabels: {
          formatter: (val, ctx) => {
            const total = values.reduce((a, b) => a + b, 0);
            return `${val} (${((val / total) * 100).toFixed(1)}%)`;
          }
        }
      }
    },
    plugins: [ChartDataLabels]
  });
}

document.getElementById("download-single-pdf").addEventListener("click", () => {
  const canvas = document.getElementById("main-chart");
  const chartTitle = chartInstance?.options?.plugins?.title?.text || "Chart";
  if (!canvas || !chartInstance) return alert("No chart available to download.");
  const imageData = canvas.toDataURL("image/png");
  const start = document.getElementById("start-date").value || "All Time";
  const end = document.getElementById("end-date").value || "Now";
  const groupBy = document.getElementById("group-select").value === "yes" ? "Yes" : "No";
  const pdf = new jsPDF("landscape", "mm", "a4");
  pdf.setFontSize(12);
  pdf.text(`Date Range: ${start} to ${end}`, 10, 15);
  pdf.text(`Count replies to email as one email: ${groupBy}`, 10, 23);
  pdf.setFontSize(16);
  pdf.text(chartTitle, 10, 35);
  pdf.addImage(imageData, "PNG", 10, 40, 225, 170);
  pdf.save(`${chartTitle.replace(/\s+/g, "_").toLowerCase()}_report.pdf`);
});
