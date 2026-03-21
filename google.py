import os
import google.generativeai as genai

# 1. Apni API Key environment se lo
API_KEY = os.getenv('GOOGLE_GEMINI_API_KEY')
if not API_KEY:
    raise ValueError("GOOGLE_GEMINI_API_KEY is missing in environment")
genai.configure(api_key=API_KEY)

# 2. Model setup with instructions for detailed answers
# Humne model ko instruction di hai ki answer kam se kam 10 lines ka ho
model = genai.GenerativeModel(
    model_name='gemini-3-flash-preview',
    system_instruction="Aap ek expert mentor hain. User jo bhi pooche, uska jawab detail mein kam se kam 10 lines mein dein."
)

def start_chat():
    print("--- Gemini CMD Bot (Type 'exit' to stop) ---")
    
    while True:
        # User se input lena
        user_query = input("\nAapka Sawal: ")
        
        if user_query.lower() in ['exit', 'quit', 'bye']:
            print("Chalo milte hain! Bye.")
            break
            
        print("\nThinking... (Detailed answer generate ho raha hai)\n")
        
        try:
            # Response generate karna
            response = model.generate_content(user_query)
            
            print("-" * 50)
            print(response.text)
            print("-" * 50)
            
        except Exception as e:
            print(f"Kuch error aaya hai: {e}")

if __name__ == "__main__":
    start_chat()