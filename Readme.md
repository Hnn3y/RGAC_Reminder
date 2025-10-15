# Royal Gem Auto Care Nigeria Limited (RGAC)

```mermaid
flowchart TD
   A([TRIGGER]) --> B[FETCH CREDS FROM ENVIRONMENT]
   B --> C[AUTHENTICATE GOOGLE SHEETS]
   C --> D[FETCH MASTER SHEET DATA]
   D --> E[PROCESS CUSTOMER]
   E --> F[COMPUTE NEXT REMINDER]
   E --> G[Alphabetically Sort Data]
   E --> H[Write to Reminder Sheet]
   E --> I[Update Reminder Fields in Master]
   E --> J[CHECK FOR DUE/OVERDUE REMINDERS]
   F --> K[(Send Email via SMTP/SendGrid)]
   G --> K
   H --> K
   I --> K
   J --> K
   K --> L[Append to Status Log]
   L --> M[Return Summary]
   M --> N([END])
```