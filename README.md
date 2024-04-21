# Use GPT models in Google Sheets

A simple script to use OpenAI's GPT models in Google Sheets. This allows you to do things like summarisation, translation, entity extraction, classification, etc.

## Installation

1. Open the Google Sheet where you want to use ChatGPT.
2. Go to `Extensions` > `Apps Script`.
3. Copy the contents of [`Code.gs`](https://github.com/nicucalcea/sheets-llm/blob/main/Code.gs) and paste it into the script editor.
4. Copy the contents of [`settings.html`](https://github.com/nicucalcea/sheets-llm/blob/main/settings.html) and paste it into a new HTML file with the same name.
5. Save both and reload the Google Sheet.
6. You should now see a new menu item called `LLM`. Click on it and then `Settings`.
7. Enter your [OpenAI API key](https://platform.openai.com/account/api-keys) and save.

## Usage

You now have access to a new `=LLM()` function in Google Sheets. You can use it like this:

```
=LLM(<input_text>, "Summarise the text", "gpt-3.5-turbo")
```

By default, the function uses the `gpt-3.5-turbo` model, but you can change it to `gpt-4` for more advanced tasks.

## To-do
- [ ] Add support for non-OpenAI models
