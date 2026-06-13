with open('app.js', 'r', encoding='utf8') as f:
    text = f.read()

start = text.find('function downloadExcelReport(studentName) {')
end = text.find('XLSX.writeFile(wb, fileName);\n}')

if start != -1 and end != -1:
    body_start = start + len('function downloadExcelReport(studentName) {')
    body_end = end + len('XLSX.writeFile(wb, fileName);\n')
    
    body = text[body_start:body_end]
    new_body = '\n    try {' + body + '    } catch(err) { alert("Excel Error: " + err.message); console.error(err); }\n'
    
    text = text[:body_start] + new_body + text[body_end:]
    with open('app.js', 'w', encoding='utf8') as f:
        f.write(text)
    print('Injected try-catch successfully.')
else:
    print('Target string not found')
