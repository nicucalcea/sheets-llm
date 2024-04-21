// Global variables
const SETTINGS_PROPERTY_STORE = 'openai-api-key';

/**
 * Adds a custom menu to the active spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LLM')
    .addItem('Settings', 'showSettingsDialog')
    .addToUi();
}

/**
 * Shows a dialog box for setting the OpenAI API key.
 */
function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'LLM Settings');
}

/**
 * Saves the user's OpenAI API key to the script properties.
 * @param {string} apiKey The OpenAI API key to save.
 * @returns {boolean} true if the API key was saved successfully, false otherwise.
 */
function saveApiKey(apiKey) {
  try {
    PropertiesService.getScriptProperties().setProperty(SETTINGS_PROPERTY_STORE, apiKey);
    return true;
  } catch (error) {
    console.error('Failed to save API key:', error);
    return false;
  }
}

/**
 * Retrieves the saved OpenAI API key from the script properties.
 * @returns {string} The saved OpenAI API key, or an empty string if none is set.
 */
function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SETTINGS_PROPERTY_STORE);
  return apiKey && apiKey !== '' ? apiKey : '';
}

/**
 * Custom function to send a request to the OpenAI API.
 * @param {string} inputText The text to send to the API.
 * @param {string} prompt The prompt to use.
 * @param {string=} model The model to use (default: 'gpt-3.5-turbo').
 * @param {string=} temperature The temperature to use, lower is more precise, higher is more creative (default: '0.1').
 * @returns {string} The response from the OpenAI API.
 * @customfunction
 */
function LLM(inputText, prompt, model = 'gpt-3.5-turbo', temperature = 0.1) {
  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('OpenAI API key not set. Please visit the "LLM > Settings" menu to set your API key.');
  }

  const systemContent = "You are a helpful assistant.";

  const options = {
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`
    },
    method: "POST",
    payload: JSON.stringify({
      "model": model,
      "messages": [
        {
          "role": "system",
          "content": systemContent
        },
        {
          "role": "user",
          "content": `${prompt}\n\n${inputText}`
        }
      ],
      "temperature": temperature
    })
  };

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
  const json = JSON.parse(response.getContentText());
  return json.choices[0].message.content;
}