#!/usr/bin/env bash
set -euo pipefail

# 用法:
# bash deploy_pages.sh https://github.com/<OWNER>/<REPO>.git

if [ "${1-}" = "" ]; then
  echo "请传入公开仓库地址，例如:"
  echo "bash deploy_pages.sh https://github.com/<OWNER>/<REPO>.git"
  exit 1
fi

REPO_URL="$1"
ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
DOCS_DIR="$ROOT_DIR/docs"

if [ ! -d "$DOCS_DIR" ]; then
  echo "未找到 docs 目录: $DOCS_DIR"
  exit 1
fi

if [ ! -d "$ROOT_DIR/.git" ]; then
  git -C "$ROOT_DIR" init
fi

if ! git -C "$ROOT_DIR" remote get-url origin >/dev/null 2>&1; then
  git -C "$ROOT_DIR" remote add origin "$REPO_URL"
else
  git -C "$ROOT_DIR" remote set-url origin "$REPO_URL"
fi

git -C "$ROOT_DIR" add docs
git -C "$ROOT_DIR" commit -m "deploy: publish docs for GitHub Pages" || true
git -C "$ROOT_DIR" branch -M main
git -C "$ROOT_DIR" push -u origin main

echo ""
echo "已推送完成。请到 GitHub 仓库设置启用 Pages："
echo "Settings -> Pages -> Build and deployment -> Source: Deploy from a branch"
echo "Branch: main /docs"
