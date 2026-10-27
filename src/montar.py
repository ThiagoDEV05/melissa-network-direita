#!/usr/bin/env python3
"""Gera site/index.html a partir de src/pagina.html, com fontes locais e envio para o servidor local."""
import os
import re

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src = open(os.path.join(BASE, "src", "pagina.html"), encoding="utf-8").read()
head, body = src.split("<!--/HEAD-->", 1)

head = head.replace('<meta name="claude:viewport-fit-sides" content="cover">\n', "")
head = re.sub(r'<link rel="preconnect"[^>]*>\n', "", head)
head = re.sub(r'<link rel="stylesheet" href="https://fonts\.googleapis\.com[^"]*">',
              '<link rel="stylesheet" href="vendor/fonts.css">', head)
body = body.replace('const API_URL = "";', 'const API_URL = "/api";')

html = ('<!doctype html>\n<html lang="pt-BR">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
        + head.strip() + "\n</head>\n<body>\n" + body.strip() + "\n</body>\n</html>\n")
assert "googleapis" not in html and "unpkg" not in html, "ainda há recurso externo"
open(os.path.join(BASE, "site", "index.html"), "w", encoding="utf-8").write(html)
print("site/index.html gerado")
