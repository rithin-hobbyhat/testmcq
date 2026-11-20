import re

# 1. Update web/index.html
with open('web/index.html', 'r', encoding='utf-8') as f:
    html = f.read()
html = html.replace('<div class="brand-title">MCQ Practice Arena</div>', '<div class="brand-title">HOBBY HAT MCQ</div>')
html = html.replace('<div class="brand-sub" id="nav-subtitle">5,700+ Questions &bull; Technical Question Bank</div>', '<div class="brand-sub" id="nav-subtitle">6,498 Questions &bull; Technical Question Bank</div>')
with open('web/index.html', 'w', encoding='utf-8') as f:
    f.write(html)
print('web/index.html updated')

# 2. Update web/manifest.json
with open('web/manifest.json', 'r', encoding='utf-8') as f:
    manifest = f.read()
manifest = manifest.replace('"name": "MCQ Practice Arena"', '"name": "HOBBY HAT MCQ"')
manifest = manifest.replace('"short_name": "MCQ Practice"', '"short_name": "HOBBY HAT MCQ"')
with open('web/manifest.json', 'w', encoding='utf-8') as f:
    f.write(manifest)
print('web/manifest.json updated')

# 3. Update root index.html
with open('index.html', 'r', encoding='utf-8') as f:
    root_html = f.read()
root_html = root_html.replace('MCQ Practice Arena - Pocket Study Edition', 'HOBBY HAT MCQ - Pocket Study Edition')
root_html = root_html.replace('Loading MCQ Practice Arena...', 'Loading HOBBY HAT MCQ...')
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(root_html)
print('root index.html updated')

# 4. Update web/sw.js
with open('web/sw.js', 'r', encoding='utf-8') as f:
    sw = f.read()
sw = re.sub(r"const CACHE_NAME = 'mcq-practice-v\d+';", "const CACHE_NAME = 'mcq-practice-v21';", sw)
with open('web/sw.js', 'w', encoding='utf-8') as f:
    f.write(sw)
print('web/sw.js updated to v21')

# 5. Update Launch_Practice_Web.bat
with open('Launch_Practice_Web.bat', 'r', encoding='utf-8') as f:
    bat = f.read()
bat = bat.replace('MCQ Practice Arena', 'HOBBY HAT MCQ')
with open('Launch_Practice_Web.bat', 'w', encoding='utf-8') as f:
    f.write(bat)
print('Launch_Practice_Web.bat updated')

# 6. Update README.md and HOW_TO_HOST.md
for doc in ['README.md', 'HOW_TO_HOST.md']:
    try:
        with open(doc, 'r', encoding='utf-8') as f:
            dtext = f.read()
        dtext = dtext.replace('MCQ Practice Arena', 'HOBBY HAT MCQ')
        dtext = dtext.replace('MCQ Practice', 'HOBBY HAT MCQ')
        with open(doc, 'w', encoding='utf-8') as f:
            f.write(dtext)
        print(f'{doc} updated')
    except Exception as e:
        print(f'{doc}: {e}')
