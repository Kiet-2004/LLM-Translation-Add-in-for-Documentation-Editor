# ================================================================
# main.py
# This file contains the FastAPI application setup, including
# endpoint definitions, the translation function, and integration
# with Google's Generative AI API.
# ================================================================

import logging  # For logging events and errors
import os  # For accessing environment variables

# Import the configuration and model mapping from config.py.
from config import CONFIG
from fastapi import FastAPI  # FastAPI framework for building APIs
from fastapi.middleware.cors import (
    CORSMiddleware,
)  # Middleware for handling Cross-Origin Resource Sharing
from google import genai  # Google's Generative AI client library
from google.genai import types  # Specific types used by the Generative AI API
from pydantic import BaseModel  # For request data validation using Pydantic

# ------------------------------------------------
# Initialize the FastAPI App
# ------------------------------------------------
app = FastAPI()
logging.info("FastAPI server started in main.py")

# ------------------------------------------------
# Configure CORS Middleware
# ------------------------------------------------
# This middleware enables Cross-Origin requests from any origin.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins (for development; restrict in production)
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ------------------------------------------------
# Define the Request Model for Translation
# ------------------------------------------------
# The TranslationRequest model defines the expected structure of requests.
class TranslationRequest(BaseModel):
    text: str  # The input text to translate
    target_lang: str  # The target language for the translation
    model: str  # The name of the translation model to use
    style: str  # The desired style for the translation


# ------------------------------------------------
# Define the Translation Function
# ------------------------------------------------
# This function builds a detailed prompt for the translation API,
# sends the request, and processes the response.
def translate_text(request: TranslationRequest):
    logging.info(f"Received translation request: {request}")

    # Construct a prompt with clear instructions for the AI model.
    prompt = f"""
    You are a highly skilled {request.style} writer and translator.  
    Your task is to **automatically detect** the language of the given text and **translate it accurately** into {request.target_lang}.  

    ### **Guidelines**:
    - Preserve the **original meaning, tone, and style** of the text.  
    - Maintain the **formatting and structure**, including **line breaks, punctuation, and special characters**.  
    - Make the translation **concise** while keeping key details.  
    - Adapt idioms, cultural expressions, and technical terms appropriately for the target language.  

    ### **Strict Output Requirement**:
    - ❌ **Do NOT** add explanations, comments, notes, or extra output.  
    - ✅ **Only return the translated sentence(s), nothing else.**  

    ### **Input Text**:
    "{request.text}"

    ### **Expected Output (Translation Only)**:
    A **clean, natural-sounding translation** in {request.target_lang}.
    """

    try:
        # Create the API client using the API key from the environment variables.
        client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))
        # Retrieve the model code from our mapping. If not found, log a warning.
        model = None
        for code, name in CONFIG["models"].items():
            if name == request.model:
                model = code
                break
        if not model:
            logging.warning(f"Invalid model name: {request.model}")
            return "Error: Invalid model name"

        # Prepare the content to be sent to the translation API.
        contents = [
            types.Content(parts=[types.Part.from_text(text=prompt)], role="user")
        ]

        # Set up the configuration for the API response to be plain text.
        generate_content_config = types.GenerateContentConfig(
            response_mime_type="text/plain"
        )

        # Call the translation API.
        response = client.models.generate_content(
            model=model, contents=contents, config=generate_content_config
        )

        # Extract and clean the translated text from the response.
        translated_text = response.text.strip().strip("\"'")
        logging.info(f"Translation successful: {translated_text}")
        return translated_text

    except Exception as e:
        # Log any exceptions during the translation process.
        logging.error(f"Error during translation: {e}")
        return "Translation error"


# ------------------------------------------------
# Define API Endpoints
# ------------------------------------------------


# Endpoint to fetch the configuration data (e.g., available languages, models, styles).
@app.api_route("/config", methods=["GET", "POST"])
async def get_config():
    logging.info("Fetching configuration")
    # Return the configuration options extracted from the CONFIG dictionary.
    sent_config = {
        "languages": [name for name in CONFIG["languages"].values()],
        "models": [name for name in CONFIG["models"].values()],
        "styles": [name for name in CONFIG["styles"].values()],
    }
    return sent_config


# Endpoint to perform translation.
@app.post("/translate")
async def translate(request: TranslationRequest):
    # Process the translation request and return the translated text.
    result = translate_text(request)
    return {"translated_text": result}


# ------------------------------------------------
# Application Entry Point
# ------------------------------------------------
# When run directly, start the FastAPI server with uvicorn.
if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
