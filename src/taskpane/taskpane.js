/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("retrieveAll").onclick = retrieveAll;
    document.getElementById("addNewCategory").onclick = addNewCategory;
    document.getElementById("removeCategory").onclick = removeCategory;
    document.getElementById("setCategory").onclick = setCategory;
    document.getElementById("retrieveCurrent").onclick = retrieveCurrent;
    document.getElementById("unassignCategory").onclick = unassignCategory;
    document.getElementById("btn-get-content").onclick = getEmailContent;
    document.getElementById("btn-get-subject").onclick = getEmailSubject;
    document.getElementById("btn-get-sender").onclick = getEmailSender;
    document.getElementById("getEmails").onclick = getEmailsByCategory;
  }
});

export async function retrieveAll() {
  retrieveCategoriesInList();
}

export async function addNewCategory() {
  addNewCategoryToList();
  retrieveCategoriesInList();
}

export async function removeCategory() {
  removeCategoryFromList();
  retrieveCategoriesInList();
}

export async function setCategory() {
  setCategoryOfCurrentEmail();
  getCategoryOfCurrentEmail();
}

export async function retrieveCurrent() {
  getCategoryOfCurrentEmail();
}

export async function unassignCategory() {
  removeCategoryFromEmail();
}

// function getSubjectOfCurrentEmail() {
//   const item = Office.context.mailbox.item;
//   // Display the subject of the current item
//   let insertAt = document.getElementById("item-subject");
//   let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
//   insertAt.appendChild(label);
//   insertAt.appendChild(document.createElement("br"));
//   insertAt.appendChild(document.createTextNode(item.subject));
//   insertAt.appendChild(document.createElement("br"));
// }

function addNewCategoryToList() {
  const masterCategoriesToAdd = [
    {
        "displayName": "New Category!",
        "color": Office.MailboxEnums.CategoryColor.Preset0
    }
  ];

  Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        window.console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
  });
}

function removeCategoryFromList() {
  const masterCategoriesToRemove = ["New Category!"];

  Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
  });
}

function retrieveCategoriesInList() {
  // Call getAsync to retrieve all available categories

  Office.context.mailbox.masterCategories.getAsync((result) => {
    
    if (result != null && result.status === Office.AsyncResultStatus.Succeeded) {
      let categories = result.value;
      let categoryDisplay = document.getElementById("item-categories");
      
      // Clear previous categories if any
      categoryDisplay.innerHTML = "";
      
      // Add a header
      let header = document.createElement("b");
      header.textContent = "Categories available:";
      categoryDisplay.appendChild(header);
      categoryDisplay.appendChild(document.createElement("br"));

      // Loop through the categories and display each one
      categories.forEach((category) => {
        let categoryNode = document.createElement("div");
        categoryNode.textContent = `Name: ${category.displayName}`;
        categoryDisplay.appendChild(categoryNode);
      });
    } else {
      if (result == null) {
        window.console.error("Result is null!!");
        return;
      }
      console.error("Failed to retrieve categories: " + result.error.message);
    }
    return;
  });
}

function getCategoryOfCurrentEmail() {
  Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    let categoryDisplay = document.getElementById("email-categories");
    categoryDisplay.innerHTML = ""; // Clear previous content

    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      // Log the error message if available
      if (asyncResult.error) {
        console.log("Action failed with error: " + asyncResult.error.message);
      } else {
        console.log("Action failed with an unknown error.");
      }
    } else {
      const categories = asyncResult.value;

      // Loop through the categories and display each one
      if (categories.length > 0) {

        // Add a header
        let header = document.createElement("b");
        header.textContent = "Assigned Categories to this Email:";
        categoryDisplay.appendChild(header);
        categoryDisplay.appendChild(document.createElement("br"));
        categories.forEach((category) => {
          let categoryNode = document.createElement("div");
          categoryNode.textContent = `Name: ${category.displayName}`;
          categoryDisplay.appendChild(categoryNode);
        });
      } else {
        // Create a proper header element
        let header = document.createElement("b");
        header.textContent = "No Categories Are Assigned to this email.";
        categoryDisplay.appendChild(header);

      }
    }
  });
}

async function setCategoryOfCurrentEmail() {
  const item = Office.context.mailbox.item;

  // Ensure the item has a conversation ID
  if (!item.conversationId) {
    console.log("This email is not part of a thread.");
    return;
  }

  const categoriesToAdd = ["Purple category"];

  // Categorize the current email
  Office.context.mailbox.item.categories.addAsync(categoriesToAdd, async function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Successfully added category to the current email.");

      // Categorize all emails in the thread
      await categorizeEmailsInThread(item.conversationId, "Purple category");
    } else {
      console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
  });
}



// function setCategoryOfCurrentEmail() {
//   const categoriesToAdd = ["Purple category"];

//   Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
//     if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
//         console.log("Successfully added categories");
//     } else {
//         console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
//     }
//   });
// }

function removeCategoryFromEmail() {
  const categoriesToRemove = ["Purple category"];

  Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        getCategoryOfCurrentEmail();
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
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

