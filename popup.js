let format = "excel"; // Definição de formato padrão
let emailData; // Armazenar os dados extraídos para download posterior

document.getElementById("extract").addEventListener("click", () => {
  document.getElementById("status").innerText = "Iniciando extração...";
  document.getElementById("progress-bar").style.width = "0%";
  document.getElementById("format-buttons").style.display = "none"; // Ocultar opções de formato durante a extração

  // Enviar mensagem para o `background.js` iniciar a extração
  chrome.runtime.sendMessage({ action: "startExtraction", format }, (response) => {
    document.getElementById("status").innerText = response.status;
  });
});

document.getElementById("csv").addEventListener("click", () => {
  format = "csv";
  updateFormatButtons();
  initiateDownload();
});

document.getElementById("excel").addEventListener("click", () => {
  format = "excel";
  updateFormatButtons();
  initiateDownload();
});

function initiateDownload() {
  chrome.runtime.sendMessage({ action: "downloadFile", format, emailData });
}

function updateFormatButtons() {
  document.getElementById("csv").style.backgroundColor = format === "csv" ? "#ff8c33" : "#ff6600";
  document.getElementById("excel").style.backgroundColor = format === "excel" ? "#ff8c33" : "#ff6600";
}

chrome.runtime.onMessage.addListener((message) => {
  if (message.action === "updateProgress") {
    updateProgressBar(message.progress);
    if (message.progress === 100) {
      document.getElementById("status").innerText = "Extração completa!";
      document.getElementById("format-buttons").style.display = "flex"; // Mostrar botões após extração
    }
  } else if (message.action === "extractionComplete") {
    emailData = message.emailData;
  }
});

function updateProgressBar(progress) {
  document.getElementById("progress-bar").style.width = `${progress}%`;
}
