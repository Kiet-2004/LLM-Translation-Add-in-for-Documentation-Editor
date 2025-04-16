<a id="readme-top"></a>

# AI Translator Add-ins  

<details>
    <summary>Table of Contents</summary>
    <ol>
        <li><a href="#overview">Overview</a></li>
        <li><a href="#features">Key Features</a></li>
        <li><a href="#getting-started">Getting Started</a></li>
        <li><a href="#integration-guides">Integration Guides</a></li>
        <li><a href="#how-to-use">How to use</a></li>
        <li><a href="#customization">Customization</a></li>
        <li><a href="#troubleshooting">Troubleshooting</a></li>
        <li><a href="#license">License</a></li>
    </ol>
</details>


## 🚀 Overview  
This project provides a **real-time AI translation tool** as an add-in for **Google Docs** and **Microsoft Word**. It allows users to translate selected text directly in their documents while preserving formatting and context.


## ✨ Key Features <a id="features"></a>
✅ **Supports multiple languages**\
✅ **AI-powered translations** (Google Gemini & more)\
✅ **Custom translation styles** (formal, casual, technical)\
✅ **User-friendly sidebar interface**\
✅ **Customizable**


## 🛠️ Getting Started <a id="getting-started"></a>
### 🔧 Backend Setup  
1. **Clone the repository**  
   ```bash
   git clone https://github.com/yourusername/your-repo.git
   cd your-repo/backend
   ```

2. **Install dependencies**  
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure the API Key**  
   - Create a `.env` file in the `backend` directory.
   - Add your API key to the `.env` file:
    ```env
    GEMINI_API_KEY="your-api-key"
    ```
    You can get the API key (for free) from the [Google AI Studio](https://aistudio.google.com/apikey).

3. **Start the FastAPI backend**  
   ```bash
   uvicorn main:app --host 0.0.0.0 --port 8000 --reload
   ```

4. *(Optional)* **Expose API using ngrok**  
   ```bash
   ngrok http 8000
   ```


## 📚 Integration Guides  <a id="integration-guides"></a>
📌 **[Google Docs Add-on - Setup Guide](./google-docs/README.md)**  
📌 **[Microsoft Word Add-in - Setup Guide](./microsoft-word/README.md)**  


## 📖 How to use  <a id="how-to-use"></a>
1. Upon accessing the add-in, users can select text within the document and hit `Refresh Selected Text` to display the selected text in the sidebar.
2. Users can then choose the desired language, model and style for translation (The language will be automatically recognized).
3. Click on `Translate` to replace the selected text with the translated text.
    - If users wish to revert the changes, they can simply hit `Ctrl+Z` to undo the translation.
4. Hit `Clear` to clear the text in the sidebar.


## 🎨 Customization  <a id="customization"></a>
### Modify the Sidebar UI
The sidebar is built with HTML, CSS and Javascript. To customize:
1. Open the corresponding files' folder.
2. Modify the HTML, CSS and Javascript as needed.
3. Save the changes and reload the add-in.

### Options
Within the `backend` folder, you can modify the `data.json` file to chance:
- **Languages**: Add or remove languages from the list.
- **Models**: Add or remove translation models.
- **Styles**: Add or remove translation styles.


## ⚠️ Troubleshooting  <a id="troubleshooting"></a>
- **No response from API?** Ensure FastAPI backend is running  
- **Errors in translation?** Check API key configuration

For more information, please check the `.log` files in the `backend` directory.

## 📝 License  <a id="license"></a>
This project is licensed under the **MIT License** – Free to use and modify.  


<p align="right">
  <a href="#readme-top">⬆️ Back to top</a>
</p>

