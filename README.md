# Perla Andina B2B — V2.3.0

Portal B2B completo hospedado no Vercel.

- Interface pública: `https://perlaandinacatalogo.vercel.app`
- Layout: carregado diretamente do `PortalB2B.html` do Apps Script para manter a mesma interface B2B
- Backend: Google Apps Script via `/api/b2b`
- Web Push: Service Worker no próprio domínio principal
- Registro Push: `/api/register` → Apps Script
- Envio Push: `/api/send` → gateway VAPID protegido

O catálogo que existia antes desta migração está preservado no branch `backup-catalogo-pre-b2b-v230`.
