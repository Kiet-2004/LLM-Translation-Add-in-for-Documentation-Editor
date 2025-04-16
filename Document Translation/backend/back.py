import os
import uvicorn
from fastapi import FastAPI, Request  # Import Request for parsing JSON
from fastapi.middleware.cors import CORSMiddleware
from google import genai
from dotenv import load_dotenv

load_dotenv()
app = FastAPI()

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def generate(lang_src: str, lang_des: str, text: str, style: str) -> str:
    prompt = f"""
    You are a highly skilled {style} writer and translator.  
    Your task is to *automatically detect* the language of the given text and *translate it accurately* from {lang_src} into {lang_des}.  

    ### *Guidelines*:
    - Preserve the *original meaning, tone, and style* of the text.  
    - Maintain the *formatting and structure*, including *line breaks, punctuation, and special characters*.  
    - Make the translation *concise* while keeping key details.  
    - Adapt idioms, cultural expressions, and technical terms appropriately for the target language.  

    ### *Strict Output Requirement*:
    - ❌ *Do NOT* add explanations, comments, notes, or extra output.  
    - ✅ *Only return the translated sentence(s), nothing else.*  

    ### *Input Text*:
    "{text}"

    ### *Expected Output Format*:
    Translated text only
    """
    
    client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=prompt,
    )
    return response.text.strip()

@app.post("/chat")
async def chat(request: Request):  # Use Request to parse JSON properly
    try:
        data = await request.json()  # Extract JSON data
        lang_src = data.get("lang_src")
        lang_des = data.get("lang_des")
        text = data.get("text")
        style = data.get("style")

        if not all([lang_src, lang_des, text, style]):  # Validate inputs
            return {"message": "Missing required fields"}

        resMsg = generate(lang_src, lang_des, text, style)
        return {"message": resMsg}

    except Exception as e:
        return {"message": f"Server internal error: {str(e)}"}

if __name__ == "__main__":
    uvicorn.run("__main__:app", host="0.0.0.0", port=8000, reload=True)
