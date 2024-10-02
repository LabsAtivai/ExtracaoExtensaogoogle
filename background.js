chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === "startExtraction") {
    startExtraction(request.format)
      .then((emailData) => {
        chrome.runtime.sendMessage({ action: "extractionComplete", emailData });
        sendResponse({ status: "Extração concluída!" });
      })
      .catch((error) =>
        sendResponse({ status: `Erro na extração: ${error.message}` })
      );
    return true; // Indicar que a resposta será enviada de forma assíncrona
  } else if (request.action === "downloadFile") {
    // Inicia o download com base no formato solicitado e nos dados extraídos
    downloadEmailData(request.format, request.emailData);
    sendResponse({ status: "Download iniciado!" });
  }
});

async function startExtraction(format) {
  const token = await getAuthToken();
  if (!token) return;

  const headers = new Headers();
  headers.append("Authorization", `Bearer ${token}`);

  const labelResponse = await fetch(
    "https://www.googleapis.com/gmail/v1/users/me/labels",
    { headers }
  );
  const labels = await labelResponse.json();

  if (!labels.labels) {
    console.error(
      "Nenhuma caixa de correio encontrada ou erro na resposta da API",
      labels
    );
    return;
  }

  let emailData = {};
  let totalEmails = 0;
  let emailsProcessed = 0;

  async function getMessagesWithPagination(labelId, labelName, pageToken = null) {
    try {
      let url = `https://www.googleapis.com/gmail/v1/users/me/messages?labelIds=${labelId}&maxResults=100`;
      if (pageToken) url += `&pageToken=${pageToken}`;

      const response = await fetch(url, { headers });
      const messages = await response.json();

      if (!messages.messages) return;

      totalEmails += messages.messages.length;

      await Promise.all(
        messages.messages.map(async (message) => {
          try {
            const msgRes = await fetch(
              `https://www.googleapis.com/gmail/v1/users/me/messages/${message.id}?format=full`,
              { headers }
            );
            const msg = await msgRes.json();

            // Verifica se msg e msg.payload estão definidos antes de acessar
            if (!msg || !msg.payload || !msg.payload.headers) return;

            const headersMsg = msg.payload.headers;
            let snippet = msg.snippet || "Sem conteúdo de texto disponível";

            const emailDetails = {
              De: headersMsg.find((header) => header.name === "From")?.value || "Desconhecido",
              Para: headersMsg.find((header) => header.name === "To")?.value || "Desconhecido",
              Assunto: headersMsg.find((header) => header.name === "Subject")?.value || "Sem Assunto",
              "Primeiras Frases": snippet,
            };

            if (!emailData[labelName]) {
              emailData[labelName] = [];
            }
            emailData[labelName].push(emailDetails);

            emailsProcessed++;
          } catch (error) {
            console.log(`Erro ao processar a mensagem ${message.id}:`, error);
          } finally {
            const progress = Math.round((emailsProcessed / totalEmails) * 100);
            chrome.runtime.sendMessage({
              action: "updateProgress",
              progress,
            });
          }
        })
      );

      if (messages.nextPageToken) {
        await getMessagesWithPagination(labelId, labelName, messages.nextPageToken);
      }
    } catch (error) {
      console.error(`Erro ao processar mensagens para a label ${labelName}`, error);
    }
  }

  await Promise.all(
    labels.labels.map(async (label) => {
      await getMessagesWithPagination(label.id, label.name);
    })
  );

  return emailData;
}

function downloadEmailData(format, emailData) {
  let blob;
  if (format === "csv") {
    const csvData = convertToCSV(emailData);
    blob = new Blob([csvData], { type: "text/csv;charset=utf-8;" });
  } else if (format === "excel") {
    const workbook = XLSX.utils.book_new();
    Object.keys(emailData).forEach((labelName) => {
      const sheet = XLSX.utils.json_to_sheet(emailData[labelName]);
      XLSX.utils.book_append_sheet(workbook, sheet, labelName);
    });
    const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    blob = new Blob([excelBuffer], { type: "application/octet-stream" });
  }

  const url = URL.createObjectURL(blob);
  chrome.downloads.download({
    url,
    filename: `emails_extracao.${format === "csv" ? "csv" : "xlsx"}`,
    saveAs: true,
  });
}

function convertToCSV(data) {
  let csv = [];
  Object.keys(data).forEach((labelName) => {
    csv.push(`Label: ${labelName}`);
    const headers = Object.keys(data[labelName][0]).join(",");
    csv.push(headers);

    data[labelName].forEach((email) => {
      const row = Object.values(email).map((v) => `"${v}"`).join(",");
      csv.push(row);
    });

    csv.push(""); // Linha em branco entre labels
  });
  return csv.join("\n");
}

async function getAuthToken() {
  return new Promise((resolve) => {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError) {
        console.error("Erro ao obter token de autenticação:", chrome.runtime.lastError);
        resolve(null);
      } else {
        resolve(token);
      }
    });
  });
}
