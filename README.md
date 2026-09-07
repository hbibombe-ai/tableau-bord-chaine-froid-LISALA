# Tableau de bord de la chaîne du froid

Application Streamlit liée au classeur de suivi hebdomadaire de l’Antenne PEV Lisala.

## Indicateurs

- Nombre et fonctionnalité des réfrigérateurs
- Équipements en panne et date de début de panne
- Disponibilité du Fridge-tag
- Maintenance préventive
- Besoins, stocks et semaines de couverture par vaccin
- Ruptures de stock par zone et aire de santé
- Réalisation des séances et supervisions
- Alertes et exports CSV

## Google Sheets privé

Partager le Google Sheet avec l’adresse `client_emailBELONGS_TO_SERVICE_ACCOUNT` du compte de service, puis configurer les secrets Streamlit :

```toml
[google_sheet]
spreadsheet_id = "IDENTIFIANT_DU_GOOGLE_SHEET"

[google_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "...@....iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```
