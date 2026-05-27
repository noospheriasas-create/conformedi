#!/bin/bash
# ConformEDI - Bootstrap initial du repo
# Usage : ./bootstrap.sh
# Pré-requis :
#   - Tu es dans le dossier `conformedi` cloné depuis GitHub (avec PAT en URL)
#   - Le dossier conformedi_bootstrap est extrait à côté

set -e

if [ ! -d ".git" ]; then
    echo "❌ Ce script doit être lancé depuis le dossier git 'conformedi' cloné depuis GitHub."
    echo "   Clone d'abord :"
    echo "   git clone https://USERNAME:TOKEN@github.com/noospheriasas-create/conformedi.git"
    echo "   cd conformedi"
    echo "   puis place le contenu de conformedi_bootstrap dans ce dossier."
    exit 1
fi

echo "🔧 Vérification de la structure..."

# Vérification des fichiers attendus
REQUIRED_FILES=(
    "README.md"
    "CHANGELOG.md"
    ".gitignore"
    "schemas/conformedi_input_schema.json"
    "schemas/conformedi_rules_engine.json"
    "mappings/conformedi_edi_mapping.xlsx"
    "scripts/evaluate_rules.js"
    "scripts/generate_mapping.py"
    "prompts/extraction_system_prompt.md"
    "personas/README.md"
    "docs/architecture_v2.md"
    "docs/decisions/001_passage_v1_v2_hybride.md"
    ".github/workflows/validate_rules.yml"
)

for f in "${REQUIRED_FILES[@]}"; do
    if [ ! -f "$f" ]; then
        echo "  ❌ Manquant : $f"
        exit 1
    fi
    echo "  ✓ $f"
done

echo ""
echo "📝 Commit et push..."

# Configuration git si nécessaire
if ! git config user.email > /dev/null; then
    git config user.email "vincent@noospheria.fr"
    git config user.name "Vincent (Noospheria)"
fi

# Initial commit
git add .
git commit -m "feat: initial structure ConformEDI v2.0.0

- Architecture hybride (IA extraction + moteur de règles)
- 62 variables atomiques, 56 variables diagnostic
- 8 blocs métier formalisés
- Mapping EDI pré-rempli (165 questions)
- Documentation architecture v2 et ADR initial
- Workflow CI de validation JSON
" || echo "  (rien à commit ou déjà fait)"

# Push
git push origin main

echo ""
echo "✅ Bootstrap terminé."
echo ""
echo "Prochaines étapes :"
echo "  1. Ouvrir le repo sur GitHub et vérifier que tous les fichiers sont bien là"
echo "  2. Lire le README.md pour confirmer l'état des lieux"
echo "  3. Ajouter une page 'Artefacts techniques' dans le doc ClickUp 'Vue d'ensemble EDI'"
echo "     avec les liens vers les fichiers de ce repo"
echo "  4. Démarrer la validation du mapping EDI (passe 1 — confidence high)"
echo ""
