import re, html, io, math
from decimal import Decimal

src = r'C:\Users\svenp\Desktop\Monografia\fuentes\institucionales\MRE_Legalizaciones-Apostilla_2026-10-05.html'
t = io.open(src, encoding='utf-8', errors='replace').read()
t = re.sub(r'(?is)<(script|style)\b.*?</\1>', ' ', t)
t = re.sub(r'(?i)<br\s*/?>|</p>|</div>|</li>|</h\d>|</tr>', '\n', t)
t = re.sub(r'<[^>]+>', ' ', t)
t = html.unescape(t)
t = re.sub(r'[ \t\xa0]+', ' ', t)

filas = []
for ln in t.split('\n'):
    ln = ln.strip()
    if re.search(r'Gs\.\s*[\d\.]+', ln) and len(ln) < 120:
        filas.append(re.sub(r'\s+', ' ', ln))

J = Decimal('117077')  # Decreto 6225/2026
out = []
out.append('TABLA ARANCELARIA PUBLICADA POR EL MRE')
out.append('Capturada de https://www.mre.gov.py/legalizaciones-apostilla/ el 5 de octubre de 2026')
out.append('')
out.append('Jornal minimo diario vigente (Decreto 6225/2026, desde 01/07/2026): Gs. ' + f'{J:,}'.replace(',', '.'))
out.append('')
out.append('--- Importes publicados ---')
out.extend('  ' + f for f in filas)
out.append('')
out.append('--- Cotejo contra la escala legal de la Ley 1030/97 ---')
out.append(f"{'coef.':>6} {'exacto':>14} {'redondeo a 50':>16} {'publicado':>12}  coincide")
casos = [(Decimal('0.5'), 58550), (Decimal(1), 117100), (Decimal(2), 234200),
         (Decimal(3), 351250), (Decimal(5), 585400), (Decimal(10), 1170800)]
ok = True
for c, pub in casos:
    ex = J * c
    ceil50 = Decimal(math.ceil(ex / 50) * 50)
    m = (ceil50 == pub)
    ok &= m
    fmt = lambda v: f'{int(v):,}'.replace(',', '.') if v == int(v) else f'{v:,}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    out.append(f"{str(c):>6} {fmt(ex):>14} {fmt(ceil50):>16} {fmt(pub):>12}  {'si' if m else 'NO'}")
out.append('')
out.append(f'Los importes publicados son el valor legal redondeado al multiplo superior de Gs. 50: {ok}')
out.append(f'Sobreprecio de la Apostilla por el redondeo: Gs. {234200 - J*2}')

dst = r'C:\Users\svenp\Desktop\Monografia\fuentes\capturas\MRE_tabla-arancelaria_2026-10-05.txt'
io.open(dst, 'w', encoding='utf-8', newline='\n').write('\n'.join(out) + '\n')
print('\n'.join(out))
