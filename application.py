#!/usr/bin/env python3
"""
patch_v2_final.py — Corrections complètes pour streamlit_app.py v2.0
Applique :
  1. Ajout de NOM_CABINET comme constante (après APP_URL)
  2. Remplacement de "CAB EDITION" / "CAB ÉDITION" par NOM_CABINET
  3. Remplacement de "EAN" par "ISBN" dans les messages visibles
  4. Renommage normaliser_codes_ean → normaliser_codes_isbn
  5. System prompt dirigeant avec redirection cabinet
  6. Suppression des blocs EC dupliqués

Usage : python3 patch_v2_final.py streamlit_app.py
Le fichier corrigé sera sauvegardé sous streamlit_app_v2_corrige.py
"""
import sys, re, ast

if len(sys.argv) < 2:
    print("Usage: python3 patch_v2_final.py <streamlit_app.py>")
    sys.exit(1)

src = sys.argv[1]
with open(src, encoding="utf-8") as f:
    code = f.read()

fixes = 0

# ── 1. Constante NOM_CABINET ──────────────────────────────────────────────────
if 'NOM_CABINET' not in code:
    old = 'APP_URL = "https://outilaccompagnementmaisonsedition-lyvgltfbwtqo4m9tdmzofu.streamlit.app/"'
    new = old + '''

# ============================================================
# IDENTITÉ DU CABINET — modifier ici pour personnaliser
# ============================================================
NOM_CABINET = "CAB ÉDITION"
SLOGAN_CABINET = "Expert-comptable · Maisons d\'édition indépendantes"
'''
    code = code.replace(old, new, 1)
    fixes += 1
    print("✅ 1. NOM_CABINET ajouté")
else:
    print("ℹ️  1. NOM_CABINET déjà présent")

# ── 2. "CAB EDITION" / "CAB ÉDITION" → NOM_CABINET ───────────────────────────
# On ne remplace QUE dans les valeurs de strings, pas dans les commentaires
# ni dans la définition de NOM_CABINET elle-même
patterns_cabinet = [
    ('"CAB EDITION"',  'NOM_CABINET'),
    ("'CAB EDITION'",  'NOM_CABINET'),
    ('"CAB ÉDITION"',  'NOM_CABINET'),
    ("'CAB ÉDITION'",  'NOM_CABINET'),
]
lines = code.split('\n')
new_lines = []
for line in lines:
    # Ignorer lignes de définition et commentaires
    stripped = line.lstrip()
    if stripped.startswith('#') or 'NOM_CABINET =' in line:
        new_lines.append(line)
        continue
    orig = line
    for old_s, new_s in patterns_cabinet:
        line = line.replace(old_s, new_s)
    if line != orig:
        fixes += 1
    new_lines.append(line)
code = '\n'.join(new_lines)
print(f"✅ 2. Occurrences CAB EDITION → NOM_CABINET")

# ── 3. EAN → ISBN dans messages visibles ─────────────────────────────────────
replacements_isbn = [
    # Messages d'erreur/warning affichés à l'écran
    ('"Aucun EAN détecté."',        '"Aucun ISBN détecté."'),
    ("'Aucun EAN détecté.'",        "'Aucun ISBN détecté.'"),
    ('"Aucun EAN/ISBN détecté."',   '"Aucun ISBN détecté."'),
    # Dans les f-strings de rapport
    ('} EAN actifs',                '} ISBN actifs'),
    ('" EAN actifs"',               '" ISBN actifs"'),
    ("' EAN actifs'",               "' ISBN actifs'"),
    # Colonne/label visible "EAN" seul
    ('"EAN"',                       '"ISBN"'),
    # Titre de fiche
    ('"Fiche EAN"',                 '"Fiche ISBN"'),
    ("'Fiche EAN'",                 "'Fiche ISBN'"),
    # Export filename
    ('Fiche_EAN',                   'Fiche_ISBN'),
]
for old_s, new_s in replacements_isbn:
    if old_s in code:
        code = code.replace(old_s, new_s)
        fixes += 1
print(f"✅ 3. EAN → ISBN dans messages visibles")

# ── 4. Renommer fonction interne normaliser_codes_ean ────────────────────────
code = code.replace('normaliser_codes_ean(', 'normaliser_codes_isbn(')
code = code.replace('def normaliser_codes_ean(', 'def normaliser_codes_isbn(')
fixes += 1
print("✅ 4. normaliser_codes_ean → normaliser_codes_isbn")

# ── 5. System prompt dirigeant avec redirection ───────────────────────────────
old_prompt = '''SYSTEM_PROMPT_DIRIGEANT = """Tu es un assistant de pilotage pour les maisons d\'édition indépendantes.
Tu aides le dirigeant à comprendre ses données de gestion : chiffre d\'affaires, taux de retour, rentabilité par titre, trésorerie.
Réponds de façon simple, claire et sans jargon comptable excessif.
Tu n\'as accès qu\'aux données de la maison d\'édition de l\'utilisateur.
Réponds toujours en français."""'''

new_prompt = '''SYSTEM_PROMPT_DIRIGEANT = f"""Tu es un assistant de pilotage mis à disposition par le cabinet {NOM_CABINET}.
Tu aides le dirigeant à comprendre ses données de gestion : chiffre d\'affaires, taux de retour, rentabilité par ISBN, trésorerie.
Réponds de façon simple, claire et sans jargon comptable excessif.
Tu n\'as accès qu\'aux données de la maison d\'édition de l\'utilisateur.

Si la question dépasse le périmètre des données disponibles (conseil stratégique, question juridique, fiscalité, comptabilité générale), réponds exactement :
"Pour cette question, je vous invite à contacter directement votre cabinet {NOM_CABINET}."

Réponds toujours en français."""'''

if 'SYSTEM_PROMPT_DIRIGEANT' in code and 'contacter directement' not in code:
    code = code.replace(old_prompt, new_prompt, 1)
    fixes += 1
    print("✅ 5. System prompt dirigeant mis à jour avec redirection cabinet")
else:
    print("ℹ️  5. System prompt dirigeant déjà correct")

# ── 6. Supprimer les blocs EC dupliqués ──────────────────────────────────────
pages_duplicates = [
    'elif role == "ec" and page == "💰 Trésorerie prévisionnelle":',
    "elif role == \"ec\" and page == \"✍️ Droits d'auteurs\":",
    'elif role == "ec" and page == "📦 Retours & Remises":',
    'elif role == "ec" and page == "📊 Synthèse financière":',
    'elif role == "ec" and page == "🤖 Assistant IA":',
    'elif role == "ec" and page == "📄 Rapport de pilotage":',
]

for sig in pages_duplicates:
    count = code.count(sig)
    if count <= 1:
        continue
    # Trouver positions
    pos1 = code.find(sig)
    pos2 = code.find(sig, pos1 + 1)
    # Supprimer du début de la 1ère occurrence jusqu'au début de la 2ème
    line_start_1 = code.rfind('\n', 0, pos1) + 1
    line_start_2 = code.rfind('\n', 0, pos2) + 1
    code = code[:line_start_1] + code[line_start_2:]
    fixes += 1
    print(f"✅ 6. Bloc dupliqué supprimé : {sig[:55]}...")

# ── 7. Vérification syntaxe ───────────────────────────────────────────────────
try:
    ast.parse(code)
    print("\n✅ Syntaxe Python valide")
except SyntaxError as e:
    print(f"\n❌ Erreur syntaxe ligne {e.lineno}: {e.msg}")
    print("   Le fichier corrigé n'a PAS été sauvegardé.")
    sys.exit(1)

# ── 8. Sauvegarde ─────────────────────────────────────────────────────────────
out = src.replace('.py', '_v2_corrige.py')
with open(out, 'w', encoding='utf-8') as f:
    f.write(code)

lines_total = code.count('\n') + 1
print(f"\n{'='*55}")
print(f"✅ Fichier corrigé : {out}")
print(f"   {fixes} corrections appliquées — {lines_total} lignes")
