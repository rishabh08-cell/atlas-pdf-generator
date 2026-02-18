# ── Base image: Node 20 on Debian (needed for LibreOffice + Playwright) ────────
FROM node:20-bookworm-slim

# ── System dependencies ─────────────────────────────────────────────────────────
RUN apt-get update && apt-get install -y \
    # LibreOffice for PPTX → PDF
    libreoffice \
    # Playwright Chromium deps
    chromium \
    ca-certificates \
    fonts-liberation \
    libappindicator3-1 \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdbus-1-3 \
    libgdk-pixbuf2.0-0 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libx11-xcb1 \
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    xdg-utils \
    wget \
    --no-install-recommends \
  && rm -rf /var/lib/apt/lists/*

# ── Set working directory ────────────────────────────────────────────────────────
WORKDIR /app

# ── Install Node deps ───────────────────────────────────────────────────────────
COPY package.json .
RUN npm install --production

# ── Install Playwright's own Chromium (used at runtime) ────────────────────────
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright
RUN npx playwright install chromium

# ── Copy app code ───────────────────────────────────────────────────────────────
COPY . .

# ── Create tmp directory ────────────────────────────────────────────────────────
RUN mkdir -p /app/tmp && chmod 777 /app/tmp

# ── Expose port ─────────────────────────────────────────────────────────────────
ENV PORT=3000
EXPOSE 3000

# ── Start ───────────────────────────────────────────────────────────────────────
CMD ["node", "server.js"]
