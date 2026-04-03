import os, json, traceback
from dotenv import load_dotenv
from google import genai

load_dotenv()
key = os.getenv("GEMINI_API_KEY", "").strip()
model = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")
print(f"Key present: {bool(key)}, model: {model}")

client = genai.Client(api_key=key)

# Test 1: Simple call
print("\n--- Test 1: basic Gemini call ---")
try:
    r = client.models.generate_content(model=model, contents="Say hello in one word.")
    print("OK:", repr(r.text[:100]))
except Exception as e:
    print("FAIL:", type(e).__name__, e)
    traceback.print_exc()

# Test 2: Parse query call
print("\n--- Test 2: parse query ---")
try:
    prompt = (
        "You are a query parser for a university chatbot.\n"
        "Feature context: Student\n\n"
        "From the user query below, extract fields and return ONLY valid JSON (no markdown fences):\n"
        '{"clean_query":"...","name":"...","section":"...","day":"...","hour":"...","reg_no":"...","phone":"...","intent":"..."}\n\n'
        "User query: i want know about kushal"
    )
    r = client.models.generate_content(model=model, contents=prompt)
    raw = r.text.strip()
    print("RAW:", repr(raw[:300]))
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    parsed = json.loads(raw.strip())
    print("PARSED:", parsed)
except Exception as e:
    print("FAIL:", type(e).__name__, e)
    traceback.print_exc()

# Test 3: Full ai_generate_response
print("\n--- Test 3: full pipeline ---")
try:
    import university_chatbot as uc
    result = uc.ai_generate_response("i want know about kushal", "Student", None)
    print("RESULT:", repr(result[:200]) if result else "None returned")
except Exception as e:
    print("FAIL:", type(e).__name__, e)
    traceback.print_exc()
