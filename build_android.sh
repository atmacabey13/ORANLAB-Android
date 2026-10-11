#!/usr/bin/env bash
set -euo pipefail
python3 -m pip install --user buildozer
buildozer android debug
