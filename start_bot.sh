#!/bin/bash
# SPPG Report Tools — Start Bot in background
# Make sure to edit bot/.env with your real token first!

BOT_DIR="$(cd "$(dirname "$0")/bot" && pwd)"
ENV_FILE="$BOT_DIR/.env"

if [ ! -f "$ENV_FILE" ]; then
    echo "❌ .env not found at $ENV_FILE"
    echo "   Create it with: echo 'SPPG_BOT_TOKEN=your_token_here' > $ENV_FILE"
    exit 1
fi

source "$ENV_FILE"

if [ -z "$SPPG_BOT_TOKEN" ]; then
    echo "❌ SPPG_BOT_TOKEN is empty in $ENV_FILE"
    exit 1
fi

export SPPG_BOT_TOKEN
cd "$BOT_DIR"
nohup python3 bot.py > /tmp/sppg_bot.log 2>&1 &
echo "✅ SPPG Bot started (PID: $!)"
echo "   Logs: tail -f /tmp/sppg_bot.log"
echo "   Stop: kill $!"
