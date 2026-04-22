#!/bin/sh
echo "Opruimen Chromium locks..."
find /app/.wwebjs_auth \( -name "SingletonLock" -o -name "SingletonSocket" -o -name "SingletonCookieService" \) | xargs -r rm -f
echo "Locks verwijderd, server starten..."
exec node server.js
