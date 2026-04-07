# ── Base image: Node 20 on Debian (needed for LibreOffice) ───────────────────
FROM node:20-bookworm-slim

# ── System dependencies (LibreOffice only — no more Chromium/Playwright) ─────
RUN apt-get update && apt-get install -y \
    # LibreOffice for PPTX → PDF
        libreoffice \
            # Fonts for proper PDF rendering
                fonts-liberation \
                    # Java runtime required by LibreOffice for PPTX→PDF conversion
                        default-jre \
                            # Java bridge for LibreOffice (needed for PPTX import filters)
                                libreoffice-java-common \
                                    && rm -rf /var/lib/apt/lists/*

                                    # ── Set JAVA_HOME so LibreOffice can locate the JRE ──────────────────────────
                                    ENV JAVA_HOME=/usr/lib/jvm/default-java
                                    ENV PATH="${JAVA_HOME}/bin:${PATH}"

                                    # ── Verify Java + LibreOffice at build time ──────────────────────────────────
                                    RUN java -version && soffice --headless --version

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
