function removeAllItems(arr, value) {
  var i = 0;
  while (i < arr.length) {
    if (arr[i] === value) {
      arr.splice(i, 1);
    } else {
      ++i;
    }
  }
  return arr;
}

function simplifyText(text){
  return text.toLowerCase()
    .trim()
    .normalize('NFD')
    .replaceAll(" ", "")
    .replaceAll(/[\u0300-\u036f]/g, "")
    .replaceAll(/[^a-z0-9\s-]/g, "");;
}

function formatDate(date){
  const dateAtt = date.split("/");

  return new Date(`${dateAtt[2]}-${dateAtt[1]}-${dateAtt[0]}T12:00:00`);
}

function fetchTrello(apiKey, apiToken, url, method, payload = {}) {
  let options = {
    'method': method,
    'muteHttpExceptions': true,
  };

  let fullUrl = `${url}${url.includes('?') ? '&' : '?'}key=${apiKey}&token=${apiToken}`;
  
  if (method !== 'GET' && Object.keys(payload).length > 0) {
    options.payload = payload;
  }

  try {
    const response = UrlFetchApp.fetch(fullUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode >= 200 && responseCode < 300) {
      return JSON.parse(responseText);
    } else {
      Logger.log(`Erro na requisição ${method} para ${fullUrl}. Código: ${responseCode}. Resposta: ${responseText}`);
      return null;
    }
  } catch (error) {
    Logger.log(`Exceção ao fazer requisição Trello: ${error}`);
    return null;
  }
}