import os, base64

env_value = os.getenv("LINKEDIN_GOOGLESHEET_API")
decoded = base64.b64decode(env_value).decode("utf-8")
print(decoded)