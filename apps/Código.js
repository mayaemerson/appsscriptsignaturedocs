let folderId = PropertiesService.getScriptProperties().getProperty('folderId');

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Assinatura Digital')
    .addItem('Assinar Documento', 'showModal')
    .addToUi();
}

function showModal() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Assinatura Digital')
    .setWidth(400)
    .setHeight(500);

  DocumentApp.getUi()
    .showModalDialog(html, 'Assinatura Digital');
}

function saveSignature(dataURL) {
  const fileName = 'assinatura_' + DocumentApp.getActiveDocument().getName() + '.png';
  const folder = DriveApp.getFolderById(folderId);

  if (folder) {
    const blob = Utilities.newBlob(Utilities.base64Decode(dataURL.split(',')[1]), 'image/png', fileName);
    folder.createFile(blob);
  } else {
    console.log('Pasta não encontrada!');
    throw new Error('Pasta não encontrada!'); // Lança um erro para ser tratado no HTML
  }
}

function signDocument(dataURL) {
  try {
    const fileName = 'assinatura_' + DocumentApp.getActiveDocument().getName() + '.png';
    const folder = DriveApp.getFolderById(folderId);

    if (!folder) {
      throw new Error('Pasta não encontrada!');
    }

    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      throw new Error('Este documento já foi assinado. Não é possível assiná-lo novamente.');
    } else {
      saveSignature(dataURL); // Salva a assinatura
      insertSignature();       // Insere a assinatura
    }

  } catch (error) {
    console.error(error); // Log do erro no Apps Script
    throw error;         // Re-lança o erro para ser tratado no HTML
  }
}



function insertSignature() {
  const fileName = 'assinatura_' + DocumentApp.getActiveDocument().getName() + '.png';
  const folder = DriveApp.getFolderById(folderId);

  if (!folder) {
      console.log('Pasta não encontrada: ' + folderId);
      return; // Sai da função se a pasta não for encontrada
  }

  const files = folder.getFilesByName(fileName);

  if (files.hasNext()) {
    const file = files.next();
    const blob = file.getBlob();

    const document = DocumentApp.getActiveDocument();
    const body = document.getBody();

    const elements = body.getParagraphs();
    for (let i = 0; i < elements.length; i++) {
      const element = elements[i];
      if (element.getText().includes('[ASSINATURA]')) {
        element.replaceText('[ASSINATURA]', '');
        body.appendImage(blob);
        return; // Sai da função após inserir a assinatura
      }
    }
    console.log('Marcador [ASSINATURA] não encontrado no documento.'); // Mensagem se o marcador não for encontrado

  } else {
    console.log('Arquivo não encontrado: ' + fileName);
  }
}


function closeModal() {
  return HtmlService.createHtmlOutput('').close();
}