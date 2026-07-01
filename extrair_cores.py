"""
Extrair paleta da logo e atualizar identidade visual
Relatórios: PPTX, PDF, HTML com logo + cores da marca
"""
from PIL import Image
import collections

BASE = 'G:/Outros computadores/NOTEBOOK1/James/projeto_relive/Relatório/Relatório de abertura-fechamento'
LOGO_PATH = f'{BASE}/logo-semfundo.png'

img = Image.open(LOGO_PATH).convert('RGBA')

# Extrair pixels não transparentes
pixels = []
for pixel in img.getdata():
    r, g, b, a = pixel
    if a > 30:
        pixels.append((r, g, b))

# Agrupar por clusters de cor (quantização 64 níveis = 4 buckets por canal)
quantized = collections.Counter()
for r, g, b in pixels:
    qr = (r // 64) * 64 + 32
    qg = (g // 64) * 64 + 32
    qb = (b // 64) * 64 + 32
    quantized[(qr, qg, qb)] += 1

# Ordenar clusters
top_clusters = quantized.most_common(10)

print('=== PALETA DE CORES DA LOGO RELIVE ===')
print()

# Identificar famílias de cor
# Cluster 1: plum/roxo escuro
# Cluster 2: wine/vinho
# Cluster 3: coral/terracota

# Pegar cores representativas
dark_plum = None
wine = None
coral = None

for (r, g, b), count in top_clusters:
    hex_c = f'#{r:02x}{g:02x}{b:02x}'
    pct = count / len(pixels) * 100
    print(f'  {hex_c} — {pct:.1f}%')

    # Classificar
    if r < 100 and b > 90 and not dark_plum:
        dark_plum = (r, g, b)
    elif r > 100 and r < 180 and g < 100 and b < 100 and not wine:
        wine = (r, g, b)
    elif r > 180 and g < 100 and b < 80 and not coral:
        coral = (r, g, b)

# Fallback se não encontrou
if not dark_plum:
    dark_plum = (78, 40, 125)
if not wine:
    wine = (148, 62, 87)
if not coral:
    coral = (214, 84, 55)

PALETA = {
    'primary': dark_plum,        # plum escuro
    'primary_hex': f'#{dark_plum[0]:02x}{dark_plum[1]:02x}{dark_plum[2]:02x}',
    'secondary': wine,           # wine/vinho
    'secondary_hex': f'#{wine[0]:02x}{wine[1]:02x}{wine[2]:02x}',
    'accent': coral,             # coral/terracota
    'accent_hex': f'#{coral[0]:02x}{coral[1]:02x}{coral[2]:02x}',
    'dark_bg': (max(0, dark_plum[0]-30), max(0, dark_plum[1]-20), max(0, dark_plum[2]-40)),
    'dark_bg_hex': f'#{max(0,dark_plum[0]-30):02x}{max(0,dark_plum[1]-20):02x}{max(0,dark_plum[2]-40):02x}',
    'light_bg': (248, 244, 252),
    'light_bg_hex': '#f8f4fc',
    'text': (60, 40, 80),
    'text_hex': '#3c2850',
    'white': '#ffffff',
    'muted': '#8e7ca3',
}

print()
print('=== PALETA FINAL ===')
for k, v in PALETA.items():
    print(f'  {k}: {v}')

# Salvar paleta
import json
with open(f'{BASE}/paleta.json', 'w') as f:
    json.dump({k: v for k, v in PALETA.items() if isinstance(v, str) or (isinstance(v, tuple) and len(v) == 3)}, f, indent=2)

print()
print('Paleta salva em paleta.json')
print(f'Primary: {PALETA["primary_hex"]}')
print(f'Secondary: {PALETA["secondary_hex"]}')
print(f'Accent: {PALETA["accent_hex"]}')
