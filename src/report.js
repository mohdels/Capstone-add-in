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
  document.getElementById("generate-btn").onclick = () => {
    const visual = document.getElementById("visual-select").value;
    const groupByConversationId = document.getElementById("group-select").value === "yes";
    const startDate = document.getElementById("start-date").value;
    const endDate = document.getElementById("end-date").value;
    updateChart(visual, groupByConversationId, startDate, endDate);
  };

  Papa.parse(rawCSV, {
    header: true,
    skipEmptyLines: true,
    complete: (result) => {
      emailData = result.data;
      console.log("📦 Total emails from CSV:", emailData.length); // Should be 100
      document.getElementById("generate-btn").click();
    }
  });
});

function updateChart(type, groupByConvo, startDate, endDate) {
  const filteredData = filterData(emailData, groupByConvo, startDate, endDate);

  console.log("📊 Filtered dataset passed to chart:", filteredData.length); // ✅ add this

  if (chartInstance) chartInstance.destroy();

  switch (type) {
    case "daily": renderEmailsPerDay(filteredData); break;
    case "category": renderEmailsByCategory(filteredData); break;
    case "assignee": renderEmailsByAssignee(filteredData); break;
    case "attachments": renderEmailsByAttachment(filteredData); break;
  }
}

function parseCleanDate(str) {
  if (!str) return new Date("invalid");

  str = str.trim().replace(/\s{2,}/g, ' ');
  const [datePart, timePart] = str.split(" ");

  if (!datePart || !timePart) return new Date("invalid");

  let [hour, minute, second = "00"] = timePart.split(":");

  // Pad hour and minute if necessary
  if (hour.length === 1) hour = "0" + hour;
  if (minute.length === 1) minute = "0" + minute;

  const cleanTime = `${hour}:${minute}:${second}`;
  const isoString = `${datePart}T${cleanTime}`;

  return new Date(isoString);
}


function filterData(data, groupByConvo, start, end) {
  const seenConvos = new Set();

  return data.filter(email => {
    const rawDate = (email.date || "").trim();
    const emailDatePart = rawDate.split(" ")[0]; // Get just 'YYYY-MM-DD'

    if (!/^\d{4}-\d{2}-\d{2}$/.test(emailDatePart)) return false; // skip bad dates

    if (start && emailDatePart < start) return false;
    if (end && emailDatePart > end) return false;

    if (groupByConvo) {
      if (seenConvos.has(email.conversationId)) return false;
      seenConvos.add(email.conversationId);
    }

    return true;
  });
}


function renderEmailsPerDay(data) {
  const countMap = {};

  data.forEach(email => {
    const dateStr = (email.date || "").trim();

    // Extract the date part safely without converting to Date object
    const day = dateStr.split(" ")[0];  // gets '2025-01-14'
    if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) return; // skip malformed dates

    countMap[day] = (countMap[day] || 0) + 1;
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

  // ✅ Move "Unassigned" to the end
  const orderedMap = {};
  Object.keys(map)
    .filter(key => key !== "Unassigned")
    .sort()
    .forEach(key => orderedMap[key] = map[key]);

  if (map["Unassigned"]) {
    orderedMap["Unassigned"] = map["Unassigned"];
  }

  renderBarChart("Emails by Assignee", orderedMap);
}


function renderEmailsByAttachment(data) {
  let withAttachments = 0;
  let withoutAttachments = 0;
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
  // Sort entries by date
  const entries = Object.entries(map).sort(
    ([a], [b]) => new Date(a) - new Date(b)
  );
  const labels = entries.map(([k]) => k);
  const values = entries.map(([_, v]) => v);

  const canvas = document.getElementById("main-chart");

  chartInstance = new Chart(canvas, {
    type: "bar",
    data: {
      labels: labels,
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
        datalabels: {
          display: false // hide labels on bars
        }
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
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: ["#4caf50", "#ff9800", "#03a9f4", "#f44336", "#9c27b0"]
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: { display: true, text: title },
        datalabels: {
          formatter: (val, ctx) => `${val} (${((val / values.reduce((a, b) => a + b, 0)) * 100).toFixed(1)}%)`
        }
      }
    },
    plugins: [ChartDataLabels]
  });
}


document.getElementById("download-single-pdf").addEventListener("click", () => {
  const canvas = document.getElementById("main-chart");
  const chartTitle = chartInstance?.options?.plugins?.title?.text || "Chart";

  if (!canvas || !chartInstance) {
    alert("No chart available to download.");
    return;
  }

  const imageData = canvas.toDataURL("image/png");

  // Get user inputs
  const start = document.getElementById("start-date").value || "All Time";
  const end = document.getElementById("end-date").value || "Now";
  const groupBy = document.getElementById("group-select").value === "yes" ? "Yes" : "No";

  const pdf = new jsPDF("landscape", "mm", "a4");

  // Metadata (no emojis, use consistent font)
  pdf.setFontSize(12);
  pdf.text(`Date Range: ${start} to ${end}`, 10, 15);
  pdf.text(`Count replies to email as one email: ${groupBy}`, 10, 23);

  // Chart Title
  pdf.setFontSize(16);
  pdf.text(chartTitle, 10, 35);

  // Add the chart with more height
  pdf.addImage(imageData, "PNG", 10, 40, 225, 170); // 🆙 Increase height from 120 → 160

  pdf.save(`${chartTitle.replace(/\s+/g, "_").toLowerCase()}_report.pdf`);
});
