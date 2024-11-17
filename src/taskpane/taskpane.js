/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here to retrieve and display categories
   */
  //getSubjectOfCurrentEmail();
  

  //addNewCategoryToList()
  //removeCategoryFromList()
  retrieveCategoriesInList();
  //setCategoryOfCurrentEmail();
  //removeCategoryFromEmail();
  getCategoryOfCurrentEmail();
}

function getSubjectOfCurrentEmail() {
  const item = Office.context.mailbox.item;
  
  // Display the subject of the current item
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}

function addNewCategoryToList() {

  const masterCategoriesToAdd = [
    {
        "displayName": "Urgent!",
        "color": Office.MailboxEnums.CategoryColor.Preset0
    }
  ];

  Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
  });
}

function removeCategoryFromList() {
  const masterCategoriesToRemove = ["abdul"];

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
      let header = document.createElement("b").appendChild(document.createTextNode("Categories available:"));
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
        console.error("Result is null!!");
        return;
      }
      console.error("Failed to retrieve categories: " + result.error.message);
    }
    return
  });
}

function getCategoryOfCurrentEmail() {
  Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
      const categories = asyncResult.value;
      console.log("Categories:");
      categories.forEach(function (item) {
        console.log("-- " + JSON.stringify(item));
      });

      let categoryDisplay = document.getElementById("email-categories");
      // Clear previous categories if any
      categoryDisplay.innerHTML = "";
      
      // Add a header
      let header = document.createElement("b").appendChild(document.createTextNode("Assigned Categories to this Email:"));
      categoryDisplay.appendChild(header);
      categoryDisplay.appendChild(document.createElement("br"));

      // Loop through the categories and display each one
      categories.forEach((category) => {
        let categoryNode = document.createElement("div");
        categoryNode.textContent = `Name: ${category.displayName}`;
        categoryDisplay.appendChild(categoryNode);
      });
    }
  });
}

function setCategoryOfCurrentEmail() {
  const categoriesToAdd = ["Red category"];

  Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
  });
}

function removeCategoryFromEmail() {
  const categoriesToRemove = ["Red category"];

  Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
}
