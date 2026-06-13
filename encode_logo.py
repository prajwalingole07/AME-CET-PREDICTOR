import base64

with open('logo.png', 'rb') as f:
    encoded = base64.b64encode(f.read()).decode('utf-8')

# save to a js file for easier injection
with open('logo_base64.js', 'w', encoding='utf-8') as f:
    f.write('const LOGO_BASE64 = "data:image/png;base64,' + encoded + '";\n')

print('Logo encoded, length:', len(encoded))
