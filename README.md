### ENVIRONMENT
---
- Operating System: Windows 11
- Language:
    - [Python 3.11.9](https://www.python.org/downloads/release/python-3119/)
    - [NodeJS 10.9.0](https://nodejs.org/en/blog/release/v10.9.0)

---
### BACKEND PREPARATION
---

- Install the necessary packages.

```bash
pip install fastapi uvicorn google-genai python-dotenv
```

- Then, simply run the `back.py` file.

---
### MICROSOFT WORDS ADD-IN
---

- Install the Yeoman generator for Office Add-ins.

```bash
npm install -g yo generator-office
```

- Create an add-in project using the Yeoman generator.

```bash
yo office
```

- When prompted, provide the following information to create your add-in project.

![image](https://learn.microsoft.com/en-us/office/dev/add-ins/images/yo-office-word.png)

- Cd to the directory that has just been generated.

- Replace every file in the `taskpane` folder (from the Yeoman generator) with the files in the `Word` folder (from my repo).

- Run the following command to start using the add-in. Remember to run this command while you are at the folder containing the `manifest.xml` file.

```bash
npm start
```

- Wait for a few minutes, then Microsoft Word will be automatically opened with the add-in for you.

- Want to stop? Simply close the Microsoft Word window and run the following command to stop the process.

```bash
npm stop
```

---
### GOOGLE DOCS ADD-IN
---

- Since you are running the API locally, you will need to host it online. 

- First download `ngrok` for Windows

```bash
choco install go 
```

- Then add your token using the following command. You can get token from [this link](https://dashboard.ngrok.com/get-started/setup/windows). Remember to login first.

```bash
ngrok config add-authtoken <YOUR_TOKEN_HERE>
```

- Then run the following command after running the `back.py` file.

```bash
ngrok http 8000
```

- Open the document from Google Docs that you want to use the add-in.

- Go to **Extension** > **Apps Script**

- Copy the content from `Code.gs` file in the `Docs` folder from my repo and paste it to the `Code.gs` file in the `Apps Script`.

- Create a file named `sidebar.html` in the `Apps Script`. Then copy the content from `sidebar.html` in the `Docs` folder from my repo and paste it to the created `sidebar.html` file.

- `Ctrl + S` to save the files content.

- Reload the document. Then go to **Extension**. The new add-in will be created below `Apps Script`. Choose it and then start using the add-in.