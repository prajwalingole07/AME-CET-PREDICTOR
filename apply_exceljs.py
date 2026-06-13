import sys

# Replace script in HTML
with open('predictor.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Replace xlsx-js-style with exceljs
html = html.replace(
    '<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>',
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js"></script>\n    <script src="logo_base64.js"></script>'
)

with open('predictor.html', 'w', encoding='utf-8') as f:
    f.write(html)

# Replace function in JS
with open('app.js', 'r', encoding='utf-8') as f:
    js = f.read()

with open('new_exceljs_func.js', 'r', encoding='utf-8') as f:
    new_func = f.read()

start = js.find('function downloadExcelReport(studentName) {')
end = js.find('function formatBranchCode(branchCode) {')

if start != -1 and end != -1:
    js = js[:start] + new_func + '\n' + js[end:]
    with open('app.js', 'w', encoding='utf-8') as f:
        f.write(js)
    print("Replaced successfully")
else:
    print("Failed to replace function bounds")
