#!/bin/bash
# SPPG Report Tools — Start Bot in background
cd "$(dirname "$0")/bot"
source .env
export SPPG_BOT_TOKEN
nohup python3 bot.py > /tmp/sppg_bot.log 2>&1 &
echo "✅ SPPG Bot started (PID: $!)"
echo "   Logs: tail -f /tmp/sppg_bot.log"
echo "   Stop: kill $!"
