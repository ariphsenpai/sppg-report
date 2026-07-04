#!/bin/bash
# SPPG Report Tools — Start Bot in Screen
cd "$(dirname "$0")/../bot"
source .env
export SPPG_BOT_TOKEN
screen -dmS sppg_bot python3 bot.py
echo "✅ SPPG Bot started in screen session 'sppg_bot'"
echo "   View: screen -r sppg_bot"
echo "   Logs: screen -S sppg_bot -X log"
