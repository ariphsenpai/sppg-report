#!/bin/bash
set -a
. /root/sppg-report/bot/.env
set +a
cd /root/sppg-report/bot
exec python3 bot.py
