# ── Base image: Node 20 on Debian (needed for LibreOffice) ───────────────────
FROM node:20-bookworm-slim

# ── System dependencies (LibreOffice only — no more Chromium/Playwright) ─────
RUN apt-get update && apt-get install -y \
    # LibreOffice for PPTX → PDF
        libreoffice \
            # Fonts for proper PDF rendering
                fonts-liberation \
                    && rm -rf /var/lib/apt/lists/*

                    # ── Disable Java in LibreOffice (JVM cannot start in Railway's container) ────
                    # The JVM hits OS thread limits (EAGAIN) and crashes; LibreOffice does not
                    # need Java for headless PPTX→PDF conversion via the Impress filter.
                    RUN mkdir -p /root/.config/libreoffice/4/user \
                        && printf '<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n <item oor:path="/org.openoffice.Office.Java"><prop oor:name="Enable" oor:op="fuse"><value>false</value></prop></item>\n</oor:items>\n' \
                            > /root/.config/libreoffice/4/user/registrymodifications.xcu

                            # ── Verify LibreOffice at build time ─────────────────────────────────────────
                            RUN soffice --headless --version

                            # ── Set working directory ─────────────────────────────────────────────────────
                            WORKDIR /app

                            # ── Install Node deps ─────────────────────────────────────────────────────────
                            COPY package.json .
                            RUN npm install --production

                            # ── Copy app code ─────────────────────────────────────────────────────────────
                            COPY . .

                            # ── Create tmp directory ──────────────────────────────────────────────────────
                            RUN mkdir -p /app/tmp && chmod 777 /app/tmp

                            # ── Expose port ───────────────────────────────────────────────────────────────
                            ENV PORT=3000
                            EXPOSE 3000

                            # ── Start ─────────────────────────────────────────────────────────────────────
                            CMD ["node", "server.js"]
