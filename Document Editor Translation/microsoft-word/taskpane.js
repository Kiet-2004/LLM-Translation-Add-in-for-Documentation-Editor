// ✅ Global base URL (update this to your actual API URL)
const base_url = "http://127.0.0.1:8000"; // Replace with your public FastAPI URL

// Ensure Office.js is ready before running functions
Office.onReady(function() {
    console.log("Office.js is ready!");
    loadOptions();
    updateSelectedText();
});

// Fetch selected text from Word
function updateSelectedText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById("inputText").value = result.value;
        } else {
            console.error("Error retrieving selected text:", result.error.message);
        }
    });
}

// Load options (languages, models, styles) from FastAPI backend
function loadOptions() {
    console.log(`${base_url}/config`);
    fetch(`${base_url}/config`, {
        method: "GET",
        headers: {
            "Content-Type": "application/json"
        }
    })
    .then(response => response.json())
    .then(options => {
        populateDropdown("targetLang", options.languages);
        populateDropdown("modelSelect", options.models);
        populateDropdown("styleSelect", options.styles);
    })
    .catch(error => {
        console.error("Error loading options:", error);
    });
}

function populateDropdown(id, items) {
    var select = document.getElementById(id);
    select.innerHTML = "";
    items.forEach(function(item) {
        var option = document.createElement("option");
        option.value = item;
        option.textContent = item;
        select.appendChild(option);
    });
}

// Perform translation request to FastAPI and insert text into Word
function doTranslate() {
    var text = document.getElementById("inputText").value;
    var targetLang = document.getElementById("targetLang").value;
    var model = document.getElementById("modelSelect").value;
    var style = document.getElementById("styleSelect").value;

    document.getElementById("loading-indicator").style.display = "flex";

    var payload = {
        text: text,
        target_lang: targetLang,
        model: model,
        style: style
    };

    console.log(payload);

    fetch(`${base_url}/translate`, { // ✅ Uses global base_url
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify(payload)
    })
    .then(response => response.json())
    .then(data => {
        document.getElementById("loading-indicator").style.display = "none";
        Office.context.document.setSelectedDataAsync(data.translated_text, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Error inserting text:", asyncResult.error.message);
            }
        });
    })
    .catch(error => {
        console.error("Translation error:", error);
        document.getElementById("loading-indicator").style.display = "none";
    });
}

// Clear text area
function clearText() {
    document.getElementById("inputText").value = "";
    document.getElementById("loading-indicator").style.display = "none";
}
