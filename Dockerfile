FROM node:18-slim

# Install Python3 and pip for PDF generation
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    python3-venv \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Node dependencies
COPY package.json package-lock.json ./
RUN npm install

# Install Python dependencies in a venv to avoid system conflicts
COPY requirements.txt ./
RUN python3 -m venv /opt/venv && \
    /opt/venv/bin/pip install --no-cache-dir -r requirements.txt

# Make venv python available as python3
ENV PATH="/opt/venv/bin:$PATH"

# Copy all application files
COPY . .

EXPOSE 3000

CMD ["node", "server.js"]
