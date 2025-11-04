function onFormSubmitHandler(e){

  const namedValues = e.namedValues;
  console.log("namedValues: ", namedValues);

  const SCRIPT_PROPS = PropertiesService.getScriptProperties();
  const apiKey = SCRIPT_PROPS.getProperty('TRELLO_API_KEY');
  const apiToken = SCRIPT_PROPS.getProperty('TRELLO_API_TOKEN');
  const apiUrl = SCRIPT_PROPS.getProperty('TRELLO_CARD_API_URL');
  const apiListId = SCRIPT_PROPS.getProperty('TRELLO_BOARD_LIST_ID');
  const boardApiUrl = SCRIPT_PROPS.getProperty('TRELLO_BOARD_API_URL');
  const boardId = SCRIPT_PROPS.getProperty('TRELLO_BOARD_ID');
  const confidenceLevelFieldId = SCRIPT_PROPS.getProperty('TRELLO_CONFIDENCE_LEVEL_FIELD_ID');
  const customFieldsApiUrl = SCRIPT_PROPS.getProperty('TRELLO_CUSTOM_FIELDS_API_URL');
  const defaultLabel = SCRIPT_PROPS.getProperty('TRELLO_DEFAULT_CARD_LABEL');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const cellTrelloId = sheet.getRange(lastRow, 16);

  const titulo                 = namedValues["Título"][0];
  const email                  = namedValues["Endereço de e-mail"][0];
  const nome                   = namedValues["Nome completo"][0];
  const estrturaArea           = namedValues["Estrutura da área"][0];
  const area                   = namedValues["Área"][0];
  const comunicado             = namedValues["Detalhes gerais"][0];
  const dtSolicitacao          = namedValues["Carimbo de data/hora"][0];
  const dtEntrega              = namedValues["Tempo"][0];
  const acao                   = namedValues["Ação"][0];
  const publicoAlvo            = namedValues["Público-alvo"][0];
  const baseDeEmails           = namedValues["Base de emails"][0].split(",");
  const motivo                 = namedValues["Benefícios"][0];
  const comportamentos         = namedValues["Comportamentos"][0];
  const nivelConfidencialidade = namedValues["Nível de confidencialidade"][0];
  const informacoesAdicionais  = namedValues["Informações adicionais"][0].split(",");

  const name = `[${lastRow}] - ${titulo}`;

  const desc = buildDesc(
    email
    , nome
    , estrturaArea
    , area
    , comunicado
    , dtSolicitacao
    , dtEntrega
    , acao
    , publicoAlvo
    , motivo
    , comportamentos
  );

  const trelloBoardLabels = getTrelloBoardLabels(apiKey, apiToken, boardApiUrl, boardId);
  let idLabels = buildLabels(trelloBoardLabels, publicoAlvo);
  !idLabels ? defaultLabel : idLabels;

  const cardId = createTrelloCard(apiListId, apiToken, apiKey, apiUrl, name, desc, idLabels);

  cellTrelloId.setValue(cardId);

  const filesPath = removeAllItems([...baseDeEmails, ...informacoesAdicionais], "");

  filesPath.forEach((filePath) => {
    let driveFile = getDriveFile(filePath);
    createCardAttachment(cardId, apiToken, apiKey, apiUrl, driveFile);
  })

  createCardChecklist(cardId, apiToken, apiKey, apiUrl, "Checklist");
  createCardChecklist(cardId, apiToken, apiKey, apiUrl, "Design");

  setCustomFieldValue(confidenceLevelFieldId, cardId, apiToken, apiKey, apiUrl, customFieldsApiUrl, nivelConfidencialidade);
}

function createTrelloCard(apiListId, apiToken, apiKey, apiUrl, name, desc, idLabels) {
  const query = {
    'method': 'POST',
    'payload': {
      'idList': apiListId,
      'key': apiKey,
      'token': apiToken,
      'name': name,
      'desc': desc,
      'idLabels': idLabels
    }
  };

  const response = UrlFetchApp.fetch(apiUrl, query);
  const responseData = JSON.parse(response.getContentText());

  return responseData["id"];
}

function createCardAttachment(cardId, apiToken, apiKey, apiUrl, file){
  try{
    const fileBlob = file["blob"];
    const fileName = file["name"];
    const fileMimeType = file["type"];

    const apiCardAttachmentUrl = `${apiUrl}/${cardId}/attachments?key=${apiKey}&token=${apiToken}`

    const query = {
      'method': 'POST',
      'payload': {
        'name': fileName,
        'mimeType': fileMimeType,
        'file': fileBlob
      }
    };

    const response = UrlFetchApp.fetch(apiCardAttachmentUrl, query);
    const responseData = JSON.parse(response.getContentText());
  }

  catch(error){
    Logger.log(`Erro ao adicionar arquivo no card: ${error}`)
  }
}

function createCardChecklist(cardId, apiToken, apiKey, apiUrl, name){
  try{
    const apiCardChecklist = `${apiUrl}/${cardId}/checklists?key=${apiKey}&token=${apiToken}`;

    const query = {
      'method': 'POST',
      'payload': {
        'name': name
      }
    };

    const response = UrlFetchApp.fetch(apiCardChecklist, query);
    const responseData = JSON.parse(response.getContentText());
  }

  catch(error){
    Logger.log(`Erro ao adicionar checklist no card: ${error}`)
  }
}

function setCustomFieldValue(customFieldId, cardId, apiToken, apiKey, apiUrl, customFieldsApiUrl, value){
  try{
    const apiCustomFieldUrl = `${apiUrl}/${cardId}/customField/${customFieldId}/item?key=${apiKey}&token=${apiToken}`
    const options = getCustomFieldOptions(customFieldId, apiToken, apiKey, customFieldsApiUrl);
    const optionKey = simplifyText(value);
    const idValue = options[optionKey];

    const query = {
      'method': 'PUT',
      'payload': {
        'idValue': idValue
      }
    };

    const response = UrlFetchApp.fetch(apiCustomFieldUrl, query);
    const responseData = JSON.parse(response.getContentText());
  }

  catch(error){
    Logger.log(`Erro ao adicionar checklist no card: ${error}`)
  }
}

function getCustomFieldOptions(customFieldId, apiToken, apiKey, customFieldsApiUrl){
  try{
    const apiCustomFieldOptionsUrl = `${customFieldsApiUrl}/${customFieldId}/options?key=${apiKey}&token=${apiToken}`;
    
    const response = UrlFetchApp.fetch(apiCustomFieldOptionsUrl);
    const responseData = JSON.parse(response.getContentText());

    const options = {};

    responseData.forEach((option) => {
      const optionKey = simplifyText(option["value"]["text"]);
      options[optionKey] = option["_id"]
    })

    return options;
  }

  catch(error){
    Logger.log(`Erro ao buscar opções para custom field: ${error}`)
  }
}

function getDriveFile(filePath){
  console.log(`getDriveFile(${filePath})`);
  let file;

  try{
    const fileId = filePath.match(/id=([a-zA-Z0-9_-]+)/)[1];

    driveFile = DriveApp.getFileById(fileId);

    const blob = driveFile.getBlob();
    const name = driveFile.getName()
    const type = blob.getContentType();

    file = {
      "blob": blob,
      "name": name,
      "type": type
    }
  }

  catch(error){
    Logger.log(`Erro ao buscar arquivo no drive: ${error}`)
    file = false;
  }

  return file;
}

function buildDesc(
  email
  , nome
  , estrturaArea
  , area
  , comunicado
  , dtSolicitacao
  , dtEntrega
  , acao
  , publicoAlvo
  , motivo
  , comportamentos
){
  const desc = `
  ***E-mail Solicitante:*** ${email}

  ***Nome Solicitante:*** ${nome}

  ***Estrutura da sua área:*** ${estrturaArea}

  ***Área:*** ${area}

  ***Comunicado:*** ${comunicado}

  ***Dt Solicitação:*** ${dtSolicitacao}

  ***Dt Entrega:*** ${dtEntrega}

  ***Ação:*** ${acao}

  ***Público-alvo:*** ${publicoAlvo}

  ***Motivo:*** ${motivo}

  ***Comportamentos corporativos:*** ${comportamentos}
  `.replaceAll("\t", "");

  return desc;
}

function buildLabels(trelloBoardLabels, labelsSelected){
  try{
    const labelsSelectedList = labelsSelected.split(",");

    let labels = [];

    labelsSelectedList.forEach((label) => {
      let labelKey = simplifyText(label);
      let idToAdd = trelloBoardLabels[labelKey] ?? trelloBoardLabels['outro'];
      labels.push(idToAdd)
    })

    return labels.join(",");
  }

  catch(error){
    Logger.log(`Erro ao encontrar os id das labels selecionadas: ${error}`);
    return false;
  }
}

function getTrelloBoardLabels(apiKey, apiToken, boardApiUrl, boardId){
  try {
    const apiUrl = `${boardApiUrl}/${boardId}/labels?key=${apiKey}&token=${apiToken}`

    const response = UrlFetchApp.fetch(apiUrl);
    const responseData = JSON.parse(response.getContentText());

    const labels = {};

    responseData.forEach((label) => {
      const labelKey = simplifyText(label["name"]);
      labels[labelKey] = label["id"]
    })

    return labels;
  }

  catch(error){
    Logger.log(`Erro ao buscar labels disponíveis: ${error}`)
  }
}