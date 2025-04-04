
document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("contact-form");
    const statusMessage = document.getElementById("status");
  
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      statusMessage.textContent = "Sending...";
  
      const formData = {
        name: document.getElementById("name").value,
        email: document.getElementById("email").value,
        phone: document.getElementById("phone").value,
        category: document.getElementById("category").value,
        submissionContext: document.getElementById("submissionContext").value,
        date: document.getElementById("date").value,
        description: document.getElementById("description").value,
      };
      
      
  
      try {
        const response = await fetch("https://localhost:3001/send-email", {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify(formData)
        });
  
        const result = await response.json();
        statusMessage.textContent = result.message;
        form.reset();
      } catch (error) {
        console.error("Error sending email:", error);
        statusMessage.textContent = "Something went wrong. Please try again.";
      }
    });
  });

  document.getElementById("back-to-taskpane").addEventListener("click", () => {
    window.location.href = "https://localhost:3000/taskpane.html";
  });
  