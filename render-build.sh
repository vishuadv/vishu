#!/usr/bin/env bash
set -e                                # stop on first error

# ----------  system packages ----------
apt-get update
apt-get install -y wget gnupg

# ----------  wkhtmltopdf 0.12.6 binary ----------
wget -q \
  https://github.com/wkhtmltopdf/wkhtmltopdf/releases/download/0.12.6-1/wkhtmltox_0.12.6-1.focal_amd64.deb \
  -O /tmp/wkhtml.deb
apt-get install -y /tmp/wkhtml.deb
rm /tmp/wkhtml.deb

# ----------  Python dependencies ----------
pip install -r requirements.txt
