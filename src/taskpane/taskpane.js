import './taskpane.css';

Office.onReady(() => {
    document.getElementById("open-form-btn").onclick = () => {
      window.open("https://localhost:3000/form.html", "_blank");
    };
  
    document.getElementById("open-report-btn").onclick = () => {
      window.open("https://localhost:3000/report.html", "_blank");
    };
  });