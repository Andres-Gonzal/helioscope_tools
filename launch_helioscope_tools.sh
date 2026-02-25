#!/usr/bin/env bash
set -euo pipefail
cd /home/panda/Apps/Helioscope_Tools

if [ -x /home/spyder-env/bin/python3 ]; then
  exec /home/spyder-env/bin/python3 helioscope_tools_gui.py
fi

exec python3 helioscope_tools_gui.py
