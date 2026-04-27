import os
import json
import google.generativeai as genai
import glob
from PIL import Image
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY:
    raise ValueError("GEMINI_API_KEY environment variable not set")
genai.configure(api_key=API_KEY)

model = genai.GenerativeModel('gemini-2.5-flash', generation_config={"response_mime_type": "application/json"})

def test_extraction():
    folder1 = "Upload UMS Attendance Screenshot-Report (Mandatory) (File responses)"
    folder2 = "Any attendance missed but you are present_ (provide ss that you are present that day in other classes) (File responses)"
    
    files2 = glob.glob(os.path.join(folder2, "*.*"))
    
    sample_file2 = None
    for f in files2:
        if f.lower().endswith(('.png', '.jpg', '.jpeg')):
            sample_file2 = f
            break
            
    print(f"Testing on {sample_file2}...")
    
    prompt = """
    You are an expert data extractor. Extract the date-wise attendance records from this screenshot.
    Format the output strictly as JSON:
    {
      "date_wise": [
        {"date": "2024-04-10", "subject": "Subject Name", "status": "Present"}
      ]
    }
    """
    try:
        img2 = Image.open(sample_file2)
        response2 = model.generate_content([prompt, img2])
        print("API call successful!")
        print(response2.text)
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_extraction()
