/* global Office */
import Papa from "papaparse";
import jsPDF from "jspdf";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run-report").onclick = generateReport;
    document.getElementById("download-pdf").onclick = generatePDF;
  }
});

async function generateReport() {
  const csvFile = document.getElementById("csv-file").files[0];

  if (!csvFile) {
    alert("Please upload a CSV file.");
    return;
  }

  // Parse the CSV file using PapaParse
  Papa.parse(csvFile, {
    header: true, // Ensures the first row is treated as headers
    skipEmptyLines: true, // Ignore empty rows
    complete: (result) => {
      const data = result.data; // Parsed data as an array of objects
      console.log("Parsed CSV Data:", data);

      // Generate the report using the parsed data
      processAndVisualizeData(data);
    },
    error: (error) => {
      console.error("Error parsing CSV:", error);
      alert("Failed to parse the CSV file. Please check its format.");
    },
  });
}

function processAndVisualizeData(data) {
  // Initialize aggregators
  const dailyEmails = {};
  const emailsByCategory = {};
  const emailsByAssignee = {};
  let emailsWithAttachments = 0;

  data.forEach((email) => {
    // Parse and validate the date
    const date = email.date ? new Date(email.date).getDate() : null;
    if (date) {
      dailyEmails[date] = (dailyEmails[date] || 0) + 1;
    }

    // Count emails by category
    const category = email.category || "Uncategorized";
    emailsByCategory[category] = (emailsByCategory[category] || 0) + 1;

    // Count emails by assignee
    const assignee = email.assignedTo || "Unassigned";
    emailsByAssignee[assignee] = (emailsByAssignee[assignee] || 0) + 1;

    // Check attachments (normalize values to handle variations like spaces, capitalization, etc.)
    const hasAttachments =
      email.hasAttachments &&
      email.hasAttachments.trim().toLowerCase() === "true";
    if (hasAttachments) {
      emailsWithAttachments++;
    }
  });

  const emailsWithoutAttachments = data.length - emailsWithAttachments;

  // Render charts
  renderChart(
    "Emails Sent Each Day",
    "bar",
    Object.keys(dailyEmails),
    Object.values(dailyEmails),
    "daily-emails-chart"
  );

  renderChart(
    "Emails by Category",
    "pie",
    Object.keys(emailsByCategory),
    Object.values(emailsByCategory),
    "emails-by-category-chart"
  );

  renderChart(
    "Emails by Assignee",
    "bar",
    Object.keys(emailsByAssignee),
    Object.values(emailsByAssignee),
    "emails-by-assignee-chart"
  );

  renderChart(
    "Emails with Attachments",
    "doughnut",
    ["With Attachments", "Without Attachments"],
    [emailsWithAttachments, emailsWithoutAttachments],
    "emails-with-attachments-chart"
  );
}

function renderChart(title, type, labels, data, canvasId) {
  const ctx = document.getElementById(canvasId).getContext("2d");

  new Chart(ctx, {
    type: type,
    data: {
      labels: labels,
      datasets: [
        {
          label: title,
          data: data,
          backgroundColor: [
            "rgba(75, 192, 192, 0.2)",
            "rgba(54, 162, 235, 0.2)",
            "rgba(255, 206, 86, 0.2)",
            "rgba(153, 102, 255, 0.2)",
            "rgba(255, 159, 64, 0.2)",
          ],
          borderColor: [
            "rgba(75, 192, 192, 1)",
            "rgba(54, 162, 235, 1)",
            "rgba(255, 206, 86, 1)",
            "rgba(153, 102, 255, 1)",
            "rgba(255, 159, 64, 1)",
          ],
          borderWidth: 1,
        },
      ],
    },
    options: {
      plugins: {
        title: {
          display: true,
          text: title,
        },
      },
      responsive: true,
    },
  });
}

function generatePDF() {
  const pdf = new jsPDF();

  // Add charts to PDF
  addChartToPDF("daily-emails-chart", "Emails Sent Each Day", pdf, 10, 10);
  addChartToPDF("emails-by-category-chart", "Emails by Category", pdf, 10, 90);
  addChartToPDF("emails-by-assignee-chart", "Emails by Assignee", pdf, 10, 170);
  addChartToPDF("emails-with-attachments-chart", "Emails with Attachments", pdf, 10, 250);

  // Save the PDF
  pdf.save("email-report.pdf");
}

function addChartToPDF(canvasId, title, pdf, x, y) {
  const canvas = document.getElementById(canvasId);
  const imageData = canvas.toDataURL("image/png");

  pdf.text(title, x, y - 5); // Add title above the chart
  pdf.addImage(imageData, "PNG", x, y, 180, 80); // Add chart image
}
