import re
with open('templates/index.html', encoding='utf-8') as f:
    c = f.read()
m = re.search(r'src="data:image/png;base64,([^"]+)"', c)
b = m.group(1)
with open('logo_b64.txt', 'w') as f:
    f.write(b)
print("DONE LEN:", len(b))
