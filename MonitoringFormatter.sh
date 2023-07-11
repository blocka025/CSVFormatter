#!/usr/bin/env bash

set -e
open -a TextEdit filename
source .venv/bin/activate
python -u MonitoringFormatter.py
deactivate