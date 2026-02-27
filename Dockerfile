# ── Base image: Node 20 on Debian (needed for LibreOffice) ───────────────────
FROM node:20-bookworm-slim

# ── System dependencies (LibreOffice only — no more Chromium/Playwright) ─────
RUN apt-get update && apt-get install -y \
    # LibreOffice for PPTX → PDF
        libreoffice \
            # Fonts for proper PDF rendering
                fonts-liberation \
                    --no-install-recommends \
                        && rm -rf /var/lib/apt/lists/*

                        # ── Set working directory ────────────────────────────────────────────────────
                        WORKDIR /app

                        # ── Install Node deps ────────────────────────────────────────────────────────
                        COPY package.json .
                        RUN npm install --production

                        # ── Copy app code ────────────────────────────────────────────────────────────
                        COPY . .

                        # ── Create tmp directory ─────────────────────────────────────────────────────
                        RUN mkdir -p /app/tmp && chmod 777 /app/tmp

                        # ── Expose port ──────────────────────────────────────────────────────────────
                        ENV PORT=3000
                        EXPOSE 3000

                        # ── Start ────────────────────────────────────────────────────────────────────
                        CMD ["node", "server.js"]
