#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Regenera os números do relatório visual a partir de um novo export do sistema.

    pip install xlrd
    python3 atualizar-relatorio.py caminho/para/relatorio.xls

Roda a análise e substitui o bloco <script id="dados"> dentro de
relatorio-vendas.html. O texto e o desenho da página não são tocados — quando
o export passar a ter data de emissão, é aqui que a série temporal entra.
"""

import json
import pathlib
import re
import sys

import importlib.util

AQUI = pathlib.Path(__file__).parent
PAGINA = AQUI / "relatorio-vendas.html"


def carregar_analise():
    spec = importlib.util.spec_from_file_location("analise", AQUI / "analise-vendas.py")
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def main():
    if len(sys.argv) < 2:
        sys.exit("uso: python3 atualizar-relatorio.py <relatorio.xls>")

    analise = carregar_analise()
    dados = analise.analisar(analise.ler(sys.argv[1]))

    html = PAGINA.read_text(encoding="utf-8")
    bloco = '<script id="dados" type="application/json">%s</script>' % json.dumps(
        dados, ensure_ascii=False, separators=(",", ":"))
    novo, trocas = re.subn(
        r'<script id="dados" type="application/json">.*?</script>',
        lambda _: bloco, html, count=1, flags=re.S)
    if not trocas:
        sys.exit("bloco de dados não encontrado em relatorio-vendas.html")
    PAGINA.write_text(novo, encoding="utf-8")

    t = dados["totais"]
    print(f"{PAGINA.name} atualizado")
    print(f"  {dados['fonte']['clientes']} clientes · {dados['fonte']['linhas_material']} linhas")
    print(f"  faturamento    R$ {t['faturamento']:,.2f}")
    print(f"  frete repassado R$ {t['frete_repassado']:,.2f}")
    print(f"  receita própria R$ {t['receita_propria']:,.2f}")
    print(f"  volume          {t['m3']:,.2f} m³  ({t['receita_propria_m3']:.2f} R$/m³)")


if __name__ == "__main__":
    main()
