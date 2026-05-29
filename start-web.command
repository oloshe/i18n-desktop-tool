#!/bin/zsh
cd "$(dirname "$0")"
node scripts/serve-web.mjs
echo ""
echo "Press any key to close..."
read -k 1
