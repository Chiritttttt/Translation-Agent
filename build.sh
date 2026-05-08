#!/usr/bin/env bash
# ═══════════════════════════════════════════════════════════
#  Translation Agent - macOS/Linux Build Script
# ═══════════════════════════════════════════════════════════
#  Usage:  ./build.sh              (build both)
#          ./build.sh onefile      (build onefile only)
#          ./build.sh onedir       (build onedir only)
#          ./build.sh clean        (clean build artifacts)
# ═══════════════════════════════════════════════════════════
set -euo pipefail

# ── Colors ──
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m'  # No Color

# ── Directories ──
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${SCRIPT_DIR}/.venv-build"
DIST_DIR="${SCRIPT_DIR}/dist"
BUILD_ONEFILE=1
BUILD_ONEDIR=1
DO_CLEAN=0

echo -e ""
echo -e "${CYAN}══════════════════════════════════════════════════${NC}"
echo -e "${CYAN}  Translation Agent - Build Script (macOS/Linux)${NC}"
echo -e "${CYAN}══════════════════════════════════════════════════${NC}"
echo -e ""

# ── Parse arguments ──
case "${1:-}" in
    onefile)
        BUILD_ONEDIR=0
        ;;
    onedir)
        BUILD_ONEFILE=0
        ;;
    clean)
        DO_CLEAN=1
        BUILD_ONEFILE=0
        BUILD_ONEDIR=0
        ;;
    "")
        ;;
    *)
        echo -e "${RED}[ERROR] Unknown argument: $1${NC}"
        echo "Usage: $0 [onefile|onedir|clean]"
        exit 1
        ;;
esac

# ── Clean mode ──
if [ "$DO_CLEAN" -eq 1 ]; then
    echo -e "${YELLOW}[CLEAN] Removing build artifacts...${NC}"
    rm -rf "${SCRIPT_DIR}/build"
    rm -rf "${DIST_DIR}"
    rm -rf "${SCRIPT_DIR}/.venv-build"
    echo -e "${GREEN}[CLEAN] Done.${NC}"
    exit 0
fi

# ── Check Python ──
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}[ERROR] python3 not found.${NC}"
    echo "Please install Python 3.9+:"
    echo "  macOS:  brew install python@3.11"
    echo "  Ubuntu: sudo apt install python3.11 python3.11-venv"
    exit 1
fi

PYVER=$(python3 --version 2>&1 | awk '{print $2}')
echo -e "${BLUE}[INFO] Python version: ${PYVER}${NC}"

# Check Python version >= 3.9
PYVER_MAJOR=$(echo "$PYVER" | cut -d. -f1)
PYVER_MINOR=$(echo "$PYVER" | cut -d. -f2)
if [ "$PYVER_MAJOR" -lt 3 ] || { [ "$PYVER_MAJOR" -eq 3 ] && [ "$PYVER_MINOR" -lt 9 ]; }; then
    echo -e "${RED}[ERROR] Python 3.9+ required, found ${PYVER}${NC}"
    exit 1
fi

# ── Create virtual environment ──
if [ ! -d "${VENV_DIR}" ]; then
    echo -e "${BLUE}[SETUP] Creating virtual environment...${NC}"
    python3 -m venv "${VENV_DIR}"
    echo -e "${GREEN}[SETUP] Virtual environment created.${NC}"
else
    echo -e "${BLUE}[SETUP] Using existing virtual environment.${NC}"
fi

# ── Activate venv ──
# shellcheck disable=SC1091
source "${VENV_DIR}/bin/activate"

echo -e "${BLUE}[SETUP] Using venv: $(which python)${NC}"

# ── Upgrade pip ──
echo -e "${BLUE}[SETUP] Upgrading pip...${NC}"
pip install --upgrade pip --quiet 2>/dev/null || echo -e "${YELLOW}[WARN] pip upgrade had issues, continuing...${NC}"

# ── Install dependencies ──
echo -e "${BLUE}[SETUP] Installing dependencies...${NC}"
pip install -r "${SCRIPT_DIR}/requirements-desktop.txt" --quiet
if [ $? -ne 0 ]; then
    echo -e "${RED}[ERROR] Failed to install dependencies.${NC}"
    echo "Try: pip install -r requirements-desktop.txt"
    exit 1
fi
echo -e "${GREEN}[SETUP] Dependencies installed.${NC}"

# ── Build onefile ──
if [ "$BUILD_ONEFILE" -eq 1 ]; then
    echo -e ""
    echo -e "${CYAN}──────────────────────────────────────${NC}"
    echo -e "${CYAN}  Building ONEFILE (single binary)...${NC}"
    echo -e "${CYAN}──────────────────────────────────────${NC}"
    pyinstaller "${SCRIPT_DIR}/TranslationAgent-onefile.spec" --clean --noconfirm
    if [ $? -ne 0 ]; then
        echo -e "${RED}[ERROR] onefile build failed!${NC}"
        exit 1
    fi
    echo -e "${GREEN}[OK] onefile build succeeded.${NC}"
fi

# ── Build onedir ──
if [ "$BUILD_ONEDIR" -eq 1 ]; then
    echo -e ""
    echo -e "${CYAN}──────────────────────────────────────${NC}"
    echo -e "${CYAN}  Building ONEDIR (directory mode)...${NC}"
    echo -e "${CYAN}──────────────────────────────────────${NC}"
    pyinstaller "${SCRIPT_DIR}/TranslationAgent-onedir.spec" --clean --noconfirm
    if [ $? -ne 0 ]; then
        echo -e "${RED}[ERROR] onedir build failed!${NC}"
        exit 1
    fi
    echo -e "${GREEN}[OK] onedir build succeeded.${NC}"
fi

# ── Copy .env.example ──
if [ -f "${SCRIPT_DIR}/.env.example" ]; then
    if [ "$BUILD_ONEFILE" -eq 1 ]; then
        cp "${SCRIPT_DIR}/.env.example" "${DIST_DIR}/.env.example"
    fi
    if [ "$BUILD_ONEDIR" -eq 1 ]; then
        cp "${SCRIPT_DIR}/.env.example" "${DIST_DIR}/TranslationAgent/.env.example"
    fi
    echo -e "${BLUE}[INFO] .env.example copied to output directory.${NC}"
fi

# ── macOS: create .app bundle if onedir ──
if [ "$(uname)" = "Darwin" ] && [ "$BUILD_ONEDIR" -eq 1 ]; then
    APP_DIR="${DIST_DIR}/TranslationAgent.app"
    if [ -d "${DIST_DIR}/TranslationAgent" ]; then
        echo -e "${BLUE}[INFO] macOS detected. To create .app bundle manually:${NC}"
        echo "  mkdir -p ${APP_DIR}/Contents/MacOS"
        echo "  cp -r ${DIST_DIR}/TranslationAgent/* ${APP_DIR}/Contents/MacOS/"
        echo "  # Then add Info.plist to ${APP_DIR}/Contents/"
    fi
fi

# ── Summary ──
echo -e ""
echo -e "${GREEN}══════════════════════════════════════════════════${NC}"
echo -e "${GREEN}  BUILD SUCCESSFUL${NC}"
echo -e "${GREEN}══════════════════════════════════════════════════${NC}"
echo -e ""

if [ "$BUILD_ONEFILE" -eq 1 ]; then
    EXE_NAME="TranslationAgent"
    if [ "$(uname)" = "Darwin" ]; then
        EXE_NAME="${EXE_NAME}"  # macOS no extension
    else
        EXE_NAME="${EXE_NAME}"
    fi
    echo -e "  [onefile]  ${DIST_DIR}/${EXE_NAME}"
    echo -e "             (single executable)"
fi

if [ "$BUILD_ONEDIR" -eq 1 ]; then
    echo -e "  [onedir]   ${DIST_DIR}/TranslationAgent/"
    echo -e "             (directory mode, binary + libs)"
fi

echo -e ""
echo -e "  To run: Copy .env to the same directory as the binary,"
echo -e "  then execute ./TranslationAgent"
echo -e ""
echo -e "  .env template:"
echo -e "    OPENAI_API_KEY=sk-your-key-here"
echo -e "    OPENAI_BASE_URL=https://api.openai.com/v1"
echo -e "    OPENAI_MODEL=gpt-4o"
echo -e ""

exit 0
