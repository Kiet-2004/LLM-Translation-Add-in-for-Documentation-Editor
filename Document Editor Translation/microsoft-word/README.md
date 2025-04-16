#  Microsoft Word Add-in - Setup Guide


## 🚀 Overview
The **Microsoft Word AI Translator Add-in** seamless integrates into Microsoft Word, providing a quick and efficient way to translate selected text using AI-powered models.


## 🔧 Installation & Setup
### 1️⃣ **Install Node.js (If Not Installed)**
MS Word uses Node.js to run the add-in. If you don't have Node.js installed, download and install it from [here](https://nodejs.org/).

### 2️⃣ **Install SSL Certificate**
MS Word requires the add-in to be served over HTTPS. To do this, we need to create a self-signed SSL certificate.
```bash
npm install -g office-addin-dev-certs
office-addin-dev-certs install
```

### 3️⃣ **Install via Microsoft Office Add-ins Development Kit**
- In VSCode (or your preferred code editor), install the extension `Microsoft Office Add-ins Development Kit`.
- Hit create a new add-in, follow its instructions, and choose Word as the target application.
- A new project will be created with the necessary files and configurations.
- Replace these files in the `src/taskpane` folder:
    - `taskpane.html`
    - `taskpane.js`
    - `taskpane.css`
- Replace the `manifest.xml` file in the root directory.

### 4️⃣ **Run the Add-in**
- In the terminal, run the following command to start the add-in server.
```bash
npm start
```

During the startup process, `Allow localhost loopback for Microsoft Edge WebView? (Y/n)` will be prompted. You can select either `Y` or `n` based on your preference, but it is not required for running the add-in.

*<u>Note</u>*: On the first startup, it may take a while to load the add-in in Word.


## 🎨 Customization
The sidebar UI can be customized by modifying the files in the `src/taskpane` folder. You can change the layout, styles, and functionality as needed.


## ⚠️ Troubleshooting
Open the terminal or the debug console to check for any errors.


## 📌 Related Guides
- **[Google Docs Add-on - Setup Guide](../google-docs/README.md)**
 - **[Main Project README](../README.md)**

