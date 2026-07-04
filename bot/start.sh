#!/bin/bash
# Start SPPG Bot
cd "$(dirname "$0")"
source .env
export SPPG_BOT_TOKEN
exec python3 bot.py
