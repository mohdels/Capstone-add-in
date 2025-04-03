/* global Office */
import Papa from "papaparse";
import jsPDF from "jspdf";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run-report").onclick = generateReport;
    document.getElementById("download-pdf").onclick = generatePDF;
    document.getElementById("open-form").onclick = () => {
      window.location.href = "https://localhost:3000/form.html";
    };
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

// Function to get email content
function getEmailContent() {
  const item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Text, function (result) {
    let contentDisplay = document.getElementById("email-content");
    contentDisplay.innerHTML = ""; // Clear previous content

    if (result.status === Office.AsyncResultStatus.Succeeded) {
      let header = document.createElement("b");
      header.textContent = "Email Content:";
      contentDisplay.appendChild(header);
      contentDisplay.appendChild(document.createElement("br"));

      let contentNode = document.createElement("div");
      contentNode.textContent = result.value;
      contentDisplay.appendChild(contentNode);
    } else {
      console.error("Error retrieving email content: " + result.error.message);
    }
  });
}

// Function to get email subject
function getEmailSubject() {
  const item = Office.context.mailbox.item;

  let subjectDisplay = document.getElementById("email-content");
  subjectDisplay.innerHTML = ""; // Clear previous content

  let header = document.createElement("b");
  header.textContent = "Email Subject:";
  subjectDisplay.appendChild(header);
  subjectDisplay.appendChild(document.createElement("br"));

  let subjectNode = document.createElement("div");
  subjectNode.textContent = item.subject || "No subject found.";
  subjectDisplay.appendChild(subjectNode);

  console.log("Email Subject:", item.subject);
}

// Function to get email sender
function getEmailSender() {
  const item = Office.context.mailbox.item;

  let senderDisplay = document.getElementById("email-content");
  senderDisplay.innerHTML = ""; // Clear previous content

  if (item.from) {
    let header = document.createElement("b");
    header.textContent = "Sender Information:";
    senderDisplay.appendChild(header);
    senderDisplay.appendChild(document.createElement("br"));

    let senderNode = document.createElement("div");
    senderNode.textContent = `From: ${item.from.displayName} <${item.from.emailAddress}>`;
    senderDisplay.appendChild(senderNode);

    console.log("Sender:", item.from.displayName, item.from.emailAddress);
  } else {
    let header = document.createElement("b");
    header.textContent = "No sender information available.";
    senderDisplay.appendChild(header);
  }
}

async function getEmailsByCategory() {
  const category = "Red category";
  const token = await getAccessToken(); // Function to retrieve the Microsoft Graph access token

  if (!token) {
    console.error("Failed to retrieve access token.");
    return;
  }

  const headers = new Headers();
  headers.append("Authorization", `Bearer ${token}`);
  headers.append("Content-Type", "application/json");

  // Query messages with a specific category
  const query = encodeURIComponent(`categories/any(c:c eq '${category}')`);
  const endpoint = `https://graph.microsoft.com/v1.0/me/messages?$filter=${query}`;

  try {
    const response = await fetch(endpoint, {
      method: "GET",
      headers: headers,
    });

    if (!response.ok) {
      console.error("Failed to fetch emails by category:", response.statusText);
      return;
    }

    const data = await response.json();
    displayEmailsByCategory(data.value); // Display retrieved emails
  } catch (error) {
    console.error("Error fetching emails by category:", error);
  }
}

function displayEmailsByCategory(emails) {
  const displayArea = document.getElementById("emails");
  displayArea.innerHTML = ""; // Clear previous results

  if (emails.length === 0) {
    displayArea.textContent = "No emails found for the specified category.";
    return;
  }

  let header = document.createElement("b");
  header.textContent = "Emails with the Specified Category:";
  displayArea.appendChild(header);
  displayArea.appendChild(document.createElement("br"));

  emails.forEach((email) => {
    let emailNode = document.createElement("div");
    emailNode.textContent = `Subject: ${email.subject}, From: ${email.from.emailAddress.name}`;
    displayArea.appendChild(emailNode);
  });
}

async function getAccessToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getAccessTokenAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        console.error("Failed to get access token:", result.error.message);
        reject(result.error);
      }
    });
  });
}

async function getEmailsInThread(conversationId) {
  const token = await getAccessToken();

  if (!token) {
    console.error("Failed to retrieve access token.");
    return [];
  }

  const headers = new Headers();
  headers.append("Authorization", `Bearer ${token}`);
  headers.append("Content-Type", "application/json");

  const endpoint = `https://graph.microsoft.com/v1.0/me/messages?$filter=conversationId eq '${conversationId}'`;

  try {
    const response = await fetch(endpoint, {
      method: "GET",
      headers: headers,
    });

    if (!response.ok) {
      console.error("Failed to fetch emails in the thread:", response.statusText);
      return [];
    }

    const data = await response.json();
    return data.value || [];
  } catch (error) {
    console.error("Error fetching emails in the thread:", error);
    return [];
  }
}


async function categorizeEmailsInThread(conversationId, category) {
  const emails = await getEmailsInThread(conversationId);

  if (emails.length === 0) {
    console.log("No emails found in the thread.");
    return;
  }

  const token = await getAccessToken();

  if (!token) {
    console.error("Failed to retrieve access token.");
    return;
  }

  const headers = new Headers();
  headers.append("Authorization", `Bearer ${token}`);
  headers.append("Content-Type", "application/json");

  // Loop through the emails and update their categories
  for (const email of emails) {
    const endpoint = `https://graph.microsoft.com/v1.0/me/messages/${email.id}`;
    const body = JSON.stringify({
      categories: [...(email.categories || []), category], // Add the category
    });

    try {
      const response = await fetch(endpoint, {
        method: "PATCH",
        headers: headers,
        body: body,
      });

      if (!response.ok) {
        console.error(`Failed to update email ${email.id}:`, response.statusText);
      } else {
        console.log(`Successfully updated email ${email.id}`);
      }
    } catch (error) {
      console.error(`Error updating email ${email.id}:`, error);
    }
  }
}

