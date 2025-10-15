# Royal Gem Auto Care Nigeria Limited (RGAC)

```mermaid
flowchart TD
   A([TRIGGER]) --> B [FETCH CREDS FROM ENVIRONMENT]
   B --> C[AUTHENTICATE GOOGLE SHEETS]
   C --> D[FETCH MASTER SHEET DATA]
   D --> E[PROCESS CUSTOMER]
   E --> F[COMPUTE NEXT REMINDER]

```