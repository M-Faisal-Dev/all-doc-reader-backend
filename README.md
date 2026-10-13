# All Doc Reader backend

Bun + Express service that accepts Word, Excel and PowerPoint files, converts them through local Gotenberg, and returns PDF bytes. Files are held in memory for the request; they are not saved by this backend. The mobile app caches successful PDFs for offline reading.

## Run

```sh
bun install
cp .env.example .env
bun run dev
```

Backend: `http://localhost:5000`. Gotenberg: `http://localhost:3005` (Docker mapping `3005:3000`).

- `GET /health` checks Gotenberg availability.
- `POST /api/v1/convert/word`: multipart file field `files`; supports `.doc` and `.docx`; returns `application/pdf`.
- `POST /api/v1/convert/excel`: multipart file field `files`; supports `.xls` and `.xlsx`; preserves normal workbook pagination and returns `application/pdf`.
- Morgan logs request method, URL, status and response time in development.
- `POST /api/v1/convert/powerpoint`: multipart file field `files`; supports `.ppt` and `.pptx`; returns a static slide PDF.
- 30 MB input/output limit, two active requests, 90-second conversion timeout, 20 requests/minute per IP.

```sh
curl --fail -F files=@sample.docx http://localhost:5000/api/v1/convert/word -o preview.pdf
bun run typecheck
bun test
```

On a physical phone, use the Mac's LAN IP, e.g. `EXPO_PUBLIC_CONVERSION_API_URL=http://192.168.1.10:5000`, and keep both devices on the same Wi-Fi. The mobile development build can infer this address from Metro's host. Android emulator can use `10.0.2.2:5000`. Web previews need their origin in `CORS_ORIGINS`.

This setup is for a trusted local development network. Before public deployment add authentication/TLS and protect Gotenberg behind this service. Install document fonts in Gotenberg for closer layout fidelity; LibreOffice may differ from Microsoft Word on complex documents.
# all-doc-reader-backend
