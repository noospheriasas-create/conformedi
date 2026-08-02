# Audit SEO — page web

Page autonome ([`seo-audit.html`](seo-audit.html)) permettant de lancer un audit SEO à partir d'une URL.
Un seul fichier HTML, sans dépendance externe : ouvrez-le dans un navigateur ou hébergez-le tel quel
(Lovable, Netlify, GitHub Pages, un bucket statique…).

## Ce que fait la page

- Champ de saisie d'URL + sélecteur de pays.
- Score de santé SEO (0–100) calculé à partir de l'autorité, du trafic, des backlinks et de la vitesse.
- Autorité de domaine, trafic organique mensuel + courbe d'évolution sur 12 mois.
- Répartition des mots-clés par position, backlinks (dofollow / nofollow), domaines référents.
- Core Web Vitals (LCP, INP, CLS, score de performance).
- Recommandations prioritaires générées automatiquement.
- Tableau des top mots-clés organiques (position, volume, CPC, difficulté, trafic estimé).

## Deux modes de fonctionnement

### 1. Mode démo (par défaut)
Aucune configuration. La page génère des données d'exemple **déterministes** (le même domaine donne toujours
le même résultat) calquées sur la structure réelle de l'API Ubersuggest. Idéal pour la démonstration et le design.

### 2. Mode connecté (données réelles)
Cliquez sur **« Source de données »** en haut à droite et renseignez l'URL d'un webhook n8n.
Le webhook est appelé en `POST` et doit renvoyer le JSON d'audit. L'endpoint est mémorisé dans le
`localStorage` du navigateur.

> ⚠️ Un navigateur ne peut pas appeler directement le serveur MCP Ubersuggest
> (`https://ubersuggest-mcp.neilpatelapi.com/mcp`) : protocole JSON-RPC/MCP, authentification par token
> et restrictions CORS. Il faut un backend (n8n, ici) qui détient le token et expose un webhook HTTP simple.

## Contrat d'échange

### Requête envoyée par la page (POST)
```json
{
  "url": "https://exemple.com",
  "domain": "exemple.com",
  "locId": 2250,
  "language": "fr"
}
```

### Réponse attendue par la page
La page accepte soit un objet `domain_overview` Ubersuggest brut, soit un objet enveloppe :

```json
{
  "overview": {
    "domain": "exemple.com",
    "domainAuthority": 48,
    "traffic": 83691,
    "organic": 59365,
    "backlinks": 1159216,
    "refDomains": 80035,
    "follow": 1011779,
    "noFollow": 147437,
    "domainTraffic": { "202601": { "searchTraffic": 374124 }, "…": {} },
    "organicKeywords": [
      { "keyword": "…", "position": 3, "volume": 390, "cpcDollars": 16.16, "sd": 73, "traffic": 209 }
    ]
  },
  "pagespeed": { "lcp": 2.4, "inp": 180, "cls": 0.05, "perf": 82 }
}
```

- `domainTraffic` : les 12 derniers mois (clé `YYYYMM`) alimentent la courbe de tendance.
- `organicKeywords` : les 12 premiers alimentent le tableau et la répartition par position.
- `pagespeed` : optionnel ; s'il est absent, la carte Core Web Vitals affiche « indisponible ».

## Câblage n8n (mode connecté)

Workflow minimal :

1. **Webhook** (POST, chemin `seo-audit`) — reçoit `{ url, domain, locId, language }`.
2. **MCP Client / HTTP Request** vers `https://ubersuggest-mcp.neilpatelapi.com/mcp` avec le token Ubersuggest :
   - outil `domain_overview` → `{ domain, language, locId }`
   - (optionnel) outil `pagespeed_audit` → `{ domain }`
3. **Set / Code** — assemble la réponse au format ci-dessus (`{ overview, pagespeed }`).
4. **Respond to Webhook** — renvoie le JSON.

Pense à activer les en-têtes CORS sur le nœud Webhook (`Access-Control-Allow-Origin`) pour autoriser
l'appel depuis la page hébergée.

### Outils Ubersuggest utiles pour enrichir l'audit
`domain_overview`, `pagespeed_audit`, `site_audit` (+ `site_audit_status` / `site_audit_results`),
`backlinks_overview`, `domain_keywords`, `domain_top_pages`, `keyword_overview`.
