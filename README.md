# ConformEDI

> Diagnostics EDI/LCB-FT automatisés pour courtiers en assurance — schémas, règles métier et workflows n8n.

ConformEDI transforme un entretien client de 8 questions ouvertes en un diagnostic structuré et en un pré-remplissage du questionnaire EDI 2026 (165 questions). L'architecture est **hybride** : extraction IA des variables atomiques (Claude Sonnet), puis moteur de règles déterministes (JSON versionnable), puis génération des réponses EDI.

## Architecture en bref

```
8 réponses libres
       │
       ▼ Claude Sonnet 4.6 (extraction tool_use)
62 variables atomiques (booléens + 4 enums)
       │
       ▼ evaluate_rules.js (déterministe, zéro IA)
56 variables de diagnostic + métriques de couverture
       │
       ▼ Mapping diagnostic → EDI + Haiku pour commentaires
165 réponses EDI structurées
       │
       ▼
Excel pré-rempli + rapport PDF + dashboard Lovable
```

Pour le détail, voir [`docs/architecture_v2.md`](docs/architecture_v2.md).

## Index des artefacts

### Schémas et règles (source de vérité produit)

| Fichier | Rôle | Version |
|---|---|---|
| [`schemas/conformedi_input_schema.json`](schemas/conformedi_input_schema.json) | Tool schema Claude pour extraire les 62 variables atomiques depuis les 8 réponses libres | v2.0 |
| [`schemas/conformedi_rules_engine.json`](schemas/conformedi_rules_engine.json) | Référentiel unifié des règles métier — calcule les 56 variables de diagnostic | v2.0 |

### Mapping EDI

| Fichier | Rôle | État |
|---|---|---|
| [`mappings/conformedi_edi_mapping.xlsx`](mappings/conformedi_edi_mapping.xlsx) | Table de correspondance 165 questions EDI → variables diagnostic + règles de conversion | En validation par Vincent |

### Scripts

| Fichier | Rôle |
|---|---|
| [`scripts/evaluate_rules.js`](scripts/evaluate_rules.js) | Évaluateur déterministe — à coller dans un Code node n8n |
| [`scripts/generate_mapping.py`](scripts/generate_mapping.py) | Génère le mapping EDI pré-rempli depuis le questionnaire Excel original |

### Documentation

| Fichier | Rôle |
|---|---|
| [`docs/architecture_v2.md`](docs/architecture_v2.md) | Architecture détaillée v2 (hybride IA + règles) |
| [`docs/decisions/`](docs/decisions/) | Architecture Decision Records (ADR) |

### Personas & prompts (à venir)

| Dossier | Rôle |
|---|---|
| [`personas/`](personas/) | 3 personas test : solo IARD, cabinet santé collective, groupe avec IA |
| [`prompts/`](prompts/) | Prompts système versionnés (extraction, commentaires) |

## État d'avancement

- [x] **Architecture v2 validée** : passage du tout-IA au pipeline hybride
- [x] **Schéma d'entrée** : 62 variables atomiques sur 8 blocs
- [x] **Moteur de règles** : 56 variables diagnostic, testé sur cas réel
- [x] **Mapping EDI pré-rempli** : 165 lignes, 45% en haute confiance
- [ ] **Validation du mapping** par Vincent (3 passes : high → medium → low/GAP)
- [ ] **Rédaction du prompt système d'extraction** (test sur 3 personas en console Anthropic)
- [ ] **Construction workflows n8n** (wf_01 à wf_06)
- [ ] **Intégration Lovable + Stripe + Resend**
- [ ] **QA end-to-end et mise en production**

## Stack

- **Front client** : Lovable (React)
- **Orchestration** : n8n (self-hosted Hostinger)
- **Base** : Supabase (Postgres + pgvector + auth)
- **IA** : Claude Sonnet 4.6 (extraction) + Haiku 4.5 (commentaires) via Anthropic API
- **Paiement** : Stripe
- **Emailing** : Resend
- **Référence** : Hub France IA 2026

## Pour reprendre le projet

1. Le contexte projet et la roadmap pilotage vivent dans [ClickUp — Vue d'ensemble EDI](https://app.clickup.com/) (lien à compléter)
2. Les artefacts techniques sont dans ce repo
3. Pour démarrer une session de travail avec Claude, commencer par lui partager :
   - Le README ci-présent
   - Le schéma actuel et le rules engine en lecture
   - L'état d'avancement (cocher dans la section ci-dessus)

## Licence

Propriété de Noospheria SAS. Tous droits réservés. Repo privé — toute redistribution ou réutilisation est interdite sans autorisation écrite.

## Contact

- **Founder / Product** : Vincent ([Noospheria SAS](https://noospheria.fr))
- **First customer** : Diagnostic ConformEDI vendu (référence interne)
