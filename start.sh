#!/bin/sh
echo "Lock bestanden verwijderen..."
find /app/.wwebjs_auth -name "Singleton*" -exec rm -f {} \; 2>/dev/null
find /app/.wwebjs_auth -name "lockfile" -exec rm -f {} \; 2>/dev/null
find /app/.wwebjs_auth -name "*.lock" -exec rm -f {} \; 2>/dev/null
echo "Klaar"
echo "Server starten..."
exec node server.js
