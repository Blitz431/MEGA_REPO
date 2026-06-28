from ollama_client import pull_text

result = pull_text("Correct the spelling and grammar of the following text. Return only the corrected text, nothing else, no explanations, no preamble: 'helo my naem is aarnav and i liek to code'")
print(result)