### ENVIRONMENT
---
- Operating System: Windows 11
- Language: [Python 3.11.9](https://www.python.org/downloads/release/python-3119/)

---
### BACKEND PREPARATION
---

- Install the necessary packages.

```bash
pip install fastapi uvicorn google-genai python-dotenv
```

- Create an `.env` file in the `backend` folder with the following content (You can your API key [here](https://aistudio.google.com/apikey)):

```env
GEMINI_API_KEY="<YOUR_TOKEN_HERE>"
```

- Then, run the `main.py` file.

- Since you are running the API locally, you will need to host it online. 

- First download `ngrok` for Windows

```bash
choco install ngrok
```

- Then add your token using the following command. You can get token from [this link](https://dashboard.ngrok.com/get-started/setup/windows). Remember to login first.

```bash
ngrok config add-authtoken <YOUR_TOKEN_HERE>
```

- Then run the following command after running the `main.py` file.

```bash
ngrok http 8000
```

---
### GOOGLE SHEETS ADD-IN
---
- Open the document from Google Sheets that you want to use the add-in.

- Go to **Extension** > **Apps Script**

- Copy the content from `Code.gs` file in the `sheets` folder from my repo and paste it to the `Code.gs` file in the `Apps Script`.

- Create a file named `sidebar.html` in the `Apps Script`. Then copy the content from `sidebar.html` in the `sheets` folder from my repo and paste it to the created `sidebar.html` file.

- `Ctrl + S` to save the files content.

- Reload the document. Then go to **Extension**. The new add-in will be created in the tool bar. Choose it and then start using the add-in.