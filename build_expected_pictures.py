from pathlib import Path
from openpyxl import load_workbook
import re

repo = Path('/Users/user/Frontend-Development-')
wb_path = repo / 'Frontend_Development_Practice_Bank.xlsx'
wb = load_workbook(wb_path)

owner = 'Nithya-sri-14'
repo_name = 'Frontend-Development-'
branch = 'main'

img_root = repo / 'expected-pictures'
img_root.mkdir(parents=True, exist_ok=True)

def slugify(text: str) -> str:
    return re.sub(r'[^a-z0-9]+', '-', text.lower()).strip('-')

def palette(level):
    if level == 'Beginner':
        return ('#0f766e', '#f0fdfa', '#d1fae5')
    if level == 'Intermediate':
        return ('#1d4ed8', '#eff6ff', '#dbeafe')
    return ('#b45309', '#fff7ed', '#fed7aa')

def labels_for(title, topic):
    t = title.lower()
    if 'portfolio' in t:
        return ['Navbar', 'Hero Intro', 'About Me', 'Projects Grid', 'Contact CTA']
    if 'form' in t or topic == 'Forms':
        return ['Form Header', 'Input Fields', 'Radio/Select Group', 'Validation Hint', 'Submit Button']
    if 'table' in t or topic == 'Tables':
        return ['Caption', 'Header Row', 'Data Rows', 'Totals Row', 'Legend/Note']
    if topic == 'Semantic HTML':
        return ['Header Landmark', 'Main Article', 'Aside Notes', 'Figure/Caption', 'Footer Info']
    if topic == 'Flexbox':
        return ['Toolbar Row', 'Aligned Cards', 'Action Cluster', 'Wrapped Chips', 'Footer Links']
    if topic == 'Grid':
        return ['Top Banner', 'Sidebar', 'Main Grid A', 'Main Grid B', 'Bottom Panel']
    if topic == 'Media Queries':
        return ['Desktop Layout', 'Tablet Shift', 'Mobile Stack', 'Breakpoint Note', 'Adaptive CTA']
    if topic == 'Responsive Design':
        return ['Responsive Header', 'Fluid Hero', 'Adaptive Cards', 'Flexible Content', 'Mobile Footer']
    if topic == 'Box Model':
        return ['Outer Margin', 'Border Frame', 'Inner Padding', 'Content Box', 'Spacing Rhythm']
    if topic == 'CSS':
        return ['Theme Header', 'Styled Card', 'Button States', 'Typography Scale', 'Color Tokens']
    return ['Header', 'Primary Section', 'Secondary Section', 'Feature Block', 'Footer']

def make_svg(path: Path, level: str, topic: str, title: str, qid: int):
    accent, bg, tint = palette(level)
    labels = labels_for(title, topic)

    blocks = [
        (80, 190, 1120, 80, labels[0]),
        (80, 285, 545, 130, labels[1]),
        (655, 285, 545, 130, labels[2]),
        (80, 430, 760, 180, labels[3]),
        (860, 430, 340, 180, labels[4]),
    ]

    # small layout variation by question id
    if qid % 3 == 1:
        blocks[1] = (80, 285, 1120, 110, labels[1])
        blocks[2] = (80, 410, 360, 200, labels[2])
        blocks[3] = (460, 410, 360, 200, labels[3])
        blocks[4] = (840, 410, 360, 200, labels[4])
    elif qid % 3 == 2:
        blocks[1] = (80, 285, 360, 325, labels[1])
        blocks[2] = (460, 285, 740, 100, labels[2])
        blocks[3] = (460, 400, 360, 210, labels[3])
        blocks[4] = (840, 400, 360, 210, labels[4])

    rects = []
    for i, (x,y,w,h,lab) in enumerate(blocks):
        fill = ['#e2e8f0', '#dbeafe', '#fef3c7', '#ede9fe', '#dcfce7'][i % 5]
        rects.append(f"<rect x='{x}' y='{y}' width='{w}' height='{h}' rx='12' fill='{fill}'/>")
        rects.append(f"<text x='{x+14}' y='{y+36}' font-family='Arial, sans-serif' font-size='26' font-weight='700' fill='#111827'>{lab}</text>")

    svg = f"""<svg xmlns='http://www.w3.org/2000/svg' width='1280' height='720' viewBox='0 0 1280 720'>
  <rect width='1280' height='720' fill='{bg}'/>
  <rect x='36' y='36' width='1208' height='648' rx='20' fill='white' stroke='{accent}' stroke-width='5'/>
  <rect x='60' y='60' width='1160' height='92' rx='12' fill='{tint}'/>
  <text x='78' y='98' font-family='Arial, sans-serif' font-size='30' font-weight='800' fill='{accent}'>{title}</text>
  <text x='78' y='130' font-family='Arial, sans-serif' font-size='20' fill='#334155'>{level} • {topic} • Exact Expected UI Preview</text>
  {''.join(rects)}
  <text x='80' y='660' font-family='Arial, sans-serif' font-size='20' fill='#0f172a'>Recreate this final look exactly in your solution.</text>
</svg>"""
    path.write_text(svg, encoding='utf-8')

for sheet in wb.sheetnames:
    ws = wb[sheet]
    for r in range(2, ws.max_row + 1):
        level = ws.cell(r, 2).value
        topic = ws.cell(r, 3).value
        title = ws.cell(r, 4).value
        qid = r - 1

        d = img_root / slugify(level) / slugify(topic)
        d.mkdir(parents=True, exist_ok=True)
        fname = f"q{qid:03d}-{slugify(title)[:80]}.svg"
        p = d / fname

        make_svg(p, level, topic, title, qid)

        rel = p.relative_to(repo).as_posix()
        ws.cell(r, 9).value = f"https://raw.githubusercontent.com/{owner}/{repo_name}/{branch}/{rel}"

wb.save(wb_path)
print('Updated workbook with exact picture links for all questions.')
