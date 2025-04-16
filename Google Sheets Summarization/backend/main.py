import os
import uvicorn
from fastapi import FastAPI, Request
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

def generate(text: str, style: str, model: str, custom: str) -> str:
    prompt = f"""
    You are a highly skilled writer specializing in {style} content.  
    Your task is to **create a concise summary** of the following text in the same language as the input.  

    ### Guidelines:
    - Capture the main points and essential information accurately.
    - Preserve the original meaning, tone, and style.
    - Maintain the overall structure and formatting, including line breaks and punctuation.
    - Retain idioms, cultural references, and technical terms as they appear in the original.
    - Ensure the summary is faithful to the text and does not introduce new information.
    - Do not include any explanations, comments, or additional text.

    ### Additional conditions:
    {custom if custom != "" else "None"}

    ### Input Text:
    "{text}"

    ### Output:
    """
    
    client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
    response = client.models.generate_content(
        model=model,
        contents=prompt,
    )
    return response.text.strip()

@app.post("/chat")
async def chat(request: Request):
    try:
        data = await request.json()
        model = data.get("model")
        text = data.get("text")
        style = data.get("style")
        custom = data.get("custom")

        if not all([model, text, style]):  # Note: This line seems incorrect in the original
            return {"message": "Missing required fields"}

        resMsg = generate(text, style, model, custom)  # Note: This line seems incorrect in the original
        return {"message": resMsg}

    except Exception as e:
        return {"message": f"Server internal error: {str(e)}"}

if __name__ == "__main__":
    uvicorn.run("__main__:app", host="0.0.0.0", port=8000, reload=True)