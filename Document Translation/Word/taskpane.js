const base_url = "http://127.0.0.1:8000";

Office.onReady(function() {
    console.log("Office.js is ready!");
});

async function translateText(text) {
    const lang_src = document.getElementById("lang_src").value;
    const lang_des = document.getElementById("lang_des").value;
    const style = document.getElementById("style").value;

    const response = await fetch(base_url + "/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ lang_src, lang_des, text, style })
    });
    const data = await response.json();
    return data.message;
}

function translateSelection() {
    Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        const translatedText = await translateText(selection.text);
        document.getElementById("translated_text").value = translatedText;
    });
}

function translateWholeDocument() {
    Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        const translatedText = await translateText(body.text);
        body.insertText(translatedText, "Replace");
    });
}

function replaceSelectedText() {
    Word.run(async (context) => {
        const selection = context.document.getSelection();
        const translatedText = document.getElementById("translated_text").value;
        selection.insertText(translatedText, "Replace");
        await context.sync();
    });
}
