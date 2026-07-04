#!/bin/bash
set -euo pipefail

# Claude Code on the web のセッション開始時に依存関係を揃える。
# ローカル実行時は何もしない（webのみ）。
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "$CLAUDE_PROJECT_DIR"

# アプリ本体の依存（Renderと同じ）
pip install -q -r requirements.txt

# 運用スクリプト用の依存
#   requests      : Instagram Graph API / GA4 取得
#   qrcode/pillow : ポスターのQR生成・PDF出力
#   python-pptx   : ポスターのPPTX出力
# Playwright本体はpreinstall済みブラウザ(/opt/pw-browsers)を使うためinstall不要
pip install -q requests qrcode pillow python-pptx playwright

echo "session-start: dependencies ready"
