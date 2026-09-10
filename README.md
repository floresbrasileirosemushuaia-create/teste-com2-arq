# Perla Andina B2B — V2.3.1

Portal B2B completo hospedado no Vercel.

- Interface: `index.html` (mesmo layout do Portal B2B Apps Script)
- Backend: Google Apps Script via `/api/b2b`
- Web Push: Service Worker no próprio domínio
- Registro Push: `/api/register` -> Apps Script
- Envio Push: `/api/send` -> gateway VAPID legado enquanto as chaves privadas permanecem protegidas no projeto Push existente

O branch `backup-catalogo-pre-b2b-v231` preserva o catálogo que existia antes da migração.
