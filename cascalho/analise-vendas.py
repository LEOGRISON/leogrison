#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Análise das vendas da cascalheira — Grupo Grison.

Lê o relatório de vendas exportado do sistema (Cliente x Produto, acumulado
desde o início da operação) e produz o JSON consolidado que alimenta o
relatório visual `relatorio-vendas.html`.

Uso:
    pip install xlrd
    python3 analise-vendas.py caminho/para/relatorio.xls > dados.json

Observações sobre a fonte:
  - O export NÃO tem coluna de data. Nenhuma série temporal é possível a
    partir dele; a visão mensal/anual depende de um novo export por período.
  - O relatório mistura a cascalheira com a venda de lotes ("Fração NN (m²)"),
    do loteamento Serra Bella. As frações são excluídas integralmente.
  - Nomes de clientes são publicados sem CPF/CNPJ (LGPD).
"""

import collections
import json
import re
import sys

import xlrd

# --- produtos vendidos a granel, em m³, e não por carga -----------------------
GRANEL = {"Cascalho Sub Base", "Cascalho Base", "Aterro de Argila"}

# --- capacidade da carga, extraída do próprio nome do produto -----------------
CAPACIDADE = {"12": 12, "06": 6}

# Cada produto tem duas faixas de preço bem separadas — o preço de pátio
# (o cliente retira com caminhão próprio) e o preço entregue (com o frete
# embutido). O limiar não pode ser único: o Aterro 12m³ sai do pátio a
# ~R$155 e entregue a ~R$380, enquanto o Sub Base 12m³ sai do pátio a ~R$252
# e entregue a ~R$450. O corte é encontrado no maior vão da distribuição de
# preços de cada produto, ignorando os extremos.
VAO_MINIMO = 1.30          # o vão só separa faixas se o preço saltar 30%
CONTRATO_FECHADO = 5.0    # preço acima de 5x a mediana é contrato, não carga
PISO_CARGA = 3.0          # preço abaixo de 1/3 do pátio é preço por m³, não por carga


def numero(txt):
    """Converte o número no formato brasileiro do relatório para float."""
    txt = str(txt).strip().replace(".", "").replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return 0.0


def sem_documento(nome):
    """Remove CPF/CNPJ e a cidade do nome do cliente."""
    nome = re.sub(r"\s*/\s*0\s*$", "", nome)
    nome = re.sub(r"\s*[-/]\s*\d[\d./-]{8,}\s*", " ", nome)
    nome = re.sub(r"\s*[-/]\s*(Palmas|Lajeado|Pium|Porto Nacional|Paraíso)[/\s]*(TO)?\s*$",
                  "", nome, flags=re.I)
    return re.sub(r"\s+", " ", nome).strip(" -/")


def tem_cnpj(nome):
    return bool(re.search(r"\d{14}", re.sub(r"\D", "", nome)))


def tipo_cliente(nome):
    """
    Classifica pelo cadastro, não pela forma de compra. Os dois eixos são
    independentes: existe construtora que retira com frota própria e pessoa
    física que quer o material entregue.
    """
    n = nome.lower()
    if "prefeitura" in n or "municí" in n or "secretaria" in n:
        return "Poder público"
    if not tem_cnpj(nome):
        return "Pessoa física"
    if re.search(r"loca[çc]|bobcat|m[áa]quinas e equipamentos", n):
        return "Locadora de máquinas"
    if re.search(r"constru|incorpora|engenharia|empreiteir|terraplan", n):
        return "Construtora"
    return "Empresa (outros)"


def limiares(linhas):
    """Encontra, para cada produto, o preço que separa pátio de entrega."""
    corte = {}
    for produto in {x["produto"] for x in linhas if x["produto"] not in GRANEL}:
        precos = sorted(x["preco_medio"] for x in linhas
                        if x["produto"] == produto and x["preco_medio"] > 0)
        if len(precos) < 4:
            continue
        meio = precos[len(precos) // 2]
        precos = [p for p in precos if p >= meio / CONTRATO_FECHADO
                  and p <= meio * CONTRATO_FECHADO]
        # maior salto relativo entre preços consecutivos, no miolo da série
        borda = max(1, len(precos) // 10)
        vaos = [(precos[i + 1] / precos[i], (precos[i] + precos[i + 1]) / 2)
                for i in range(borda, len(precos) - borda - 1)]
        if vaos:
            razao, ponto = max(vaos)
            if razao >= VAO_MINIMO:
                corte[produto] = ponto
    return corte


def cadastrado_por_m3(linha, patio):
    """
    Linha de produto-carga cujo preço unitário é baixo demais para ser uma
    carga: foi lançada com preço por m³ e o cadastro do produto ficou errado.
    O caso real é a Prefeitura de Palmas, 112 unidades de "Aterro 12m³" a
    R$ 14,50 — são 112 m³, não 112 cargas de 12 m³.
    """
    base = patio.get(linha["produto"])
    return bool(base) and 0 < linha["preco_medio"] < base / PISO_CARGA


def modalidade(linha, corte, patio=None):
    if linha["produto"] in GRANEL:
        return "Granel"
    patio = patio or {}
    base = patio.get(linha["produto"])
    if base and linha["preco_medio"] > base * CONTRATO_FECHADO:
        return "Contrato fechado"
    if cadastrado_por_m3(linha, patio):
        return "Granel"
    limiar = corte.get(linha["produto"])
    if limiar is None:                       # produto sem série suficiente
        return "Entrega" if linha["preco_medio"] > 200 else "Retirada"
    return "Entrega" if linha["preco_medio"] > limiar else "Retirada"


def referencias_patio(linhas, corte):
    """Preço mediano de pátio por produto — a base para estimar o frete."""
    ref = {}
    for produto in corte:
        patio = sorted(x["preco_medio"] for x in linhas
                       if x["produto"] == produto
                       and x["preco_medio"] <= corte[produto]
                       and x["preco_medio"] > corte[produto] / CONTRATO_FECHADO)
        if patio:
            ref[produto] = patio[len(patio) // 2]
    return ref


def volume_m3(linha, patio=None):
    """Volume em m³. Granel já vem em m³; carga é multiplicada pela capacidade."""
    if linha["produto"] in GRANEL or cadastrado_por_m3(linha, patio or {}):
        return linha["qtde"]
    m = re.search(r"(\d{2})m³", linha["produto"])
    return linha["qtde"] * CAPACIDADE.get(m.group(1), 12) if m else linha["qtde"] * 12


def ler(caminho):
    livro = xlrd.open_workbook(caminho, ignore_workbook_corruption=True)
    aba = livro.sheet_by_index(0)
    linhas, cliente = [], None
    for r in range(1, aba.nrows):
        celulas = [str(c.value).strip() for c in aba.row(r)]
        if celulas[0] == "Totais":
            continue
        if celulas[0] and celulas[0] != "\xa0":
            cliente = celulas[0]           # linha-mãe: só repete o total do cliente
            continue
        linhas.append({
            "cliente": cliente,
            "produto": celulas[1],
            "qtde": numero(celulas[3]),
            "qtde_faturada": numero(celulas[4]),
            "preco_medio": numero(celulas[5]),
            "valor_faturado": numero(celulas[7]),
            "total": numero(celulas[11]),
        })
    return linhas


def analisar(linhas):
    fracoes = [x for x in linhas if x["produto"].startswith("Fração")]
    mat = [x for x in linhas if not x["produto"].startswith("Fração")]
    corte = limiares(mat)
    patio = referencias_patio(mat, corte)
    for x in mat:
        x["tipo"] = tipo_cliente(x["cliente"])
        x["modalidade"] = modalidade(x, corte, patio)
        x["m3"] = volume_m3(x, patio)
        # O frete cobrado na venda entregue é repassado inteiro ao caminhoneiro:
        # não é receita da cascalheira. Estima-se pelo quanto o preço entregue
        # excede o preço de pátio do mesmo produto.
        base = patio.get(x["produto"])
        x["frete"] = (max(0.0, x["preco_medio"] - base) * x["qtde"]
                      if x["modalidade"] == "Entrega" and base else 0.0)
        x["receita_propria"] = x["total"] - x["frete"]

    faturamento = sum(x["total"] for x in mat)
    frete = sum(x["frete"] for x in mat)
    volume = sum(x["m3"] for x in mat)

    def resumir(chave):
        d = collections.defaultdict(lambda: {"faturamento": 0.0, "m3": 0.0, "frete": 0.0,
                                             "linhas": 0, "clientes": set()})
        for x in mat:
            a = d[chave(x)]
            a["faturamento"] += x["total"]
            a["m3"] += x["m3"]
            a["frete"] += x["frete"]
            a["linhas"] += 1
            a["clientes"].add(x["cliente"])
        return {k: {"faturamento": round(v["faturamento"], 2),
                    "frete_repassado": round(v["frete"], 2),
                    "receita_propria": round(v["faturamento"] - v["frete"], 2),
                    "m3": round(v["m3"], 2),
                    "linhas": v["linhas"],
                    "clientes": len(v["clientes"]),
                    "preco_m3": round(v["faturamento"] / v["m3"], 2) if v["m3"] else 0,
                    "receita_propria_m3": round((v["faturamento"] - v["frete"]) / v["m3"], 2) if v["m3"] else 0}
                for k, v in sorted(d.items(), key=lambda t: -t[1]["faturamento"])}

    # cruzamento dos dois eixos
    cruz = collections.defaultdict(float)
    cruz_m3 = collections.defaultdict(float)
    cruz_liq = collections.defaultdict(float)
    for x in mat:
        cruz[(x["tipo"], x["modalidade"])] += x["total"]
        cruz_m3[(x["tipo"], x["modalidade"])] += x["m3"]
        cruz_liq[(x["tipo"], x["modalidade"])] += x["receita_propria"]

    # ranking e concentração de clientes
    porcli = collections.defaultdict(lambda: {"faturamento": 0.0, "m3": 0.0,
                                              "modalidades": set()})
    for x in mat:
        a = porcli[x["cliente"]]
        a["faturamento"] += x["total"]
        a["m3"] += x["m3"]
        a["modalidades"].add(x["modalidade"])
    ranking = sorted(porcli.items(), key=lambda t: -t[1]["faturamento"])
    acumulado, pareto = 0.0, []
    for i, (_, v) in enumerate(ranking, 1):
        acumulado += v["faturamento"]
        pareto.append({"n": i, "pct": round(acumulado / faturamento * 100, 2)})

    # recorrência: quantos clientes compram uma vez e quantos voltam
    faixas = collections.Counter()
    for _, v in ranking:
        c = v["m3"] / 12          # cargas equivalentes de 12m³
        faixa = ("1 carga" if c <= 1.5 else "2 a 5" if c <= 5 else
                 "6 a 20" if c <= 20 else "21 a 50" if c <= 50 else "mais de 50")
        faixas[faixa] += 1

    return {
        "fonte": {
            "linhas_material": len(mat),
            "clientes": len(porcli),
            "fracoes_excluidas": {"linhas": len(fracoes),
                                  "valor": round(sum(x["total"] for x in fracoes), 2)},
            "sem_data": True,
        },
        "precos": {
            "limiar_entrega": {k: round(v, 2) for k, v in sorted(corte.items())},
            "patio": {k: round(v, 2) for k, v in sorted(patio.items())},
        },
        "totais": {
            "faturamento": round(faturamento, 2),
            "frete_repassado": round(frete, 2),
            "receita_propria": round(faturamento - frete, 2),
            "m3": round(volume, 2),
            "preco_medio_m3": round(faturamento / volume, 2),
            "receita_propria_m3": round((faturamento - frete) / volume, 2),
            "ticket_medio_cliente": round(faturamento / len(porcli), 2),
            "valor_faturado_nf": round(sum(x["valor_faturado"] for x in mat), 2),
        },
        "por_produto": resumir(lambda x: x["produto"]),
        "por_modalidade": resumir(lambda x: x["modalidade"]),
        "por_tipo": resumir(lambda x: x["tipo"]),
        "cruzamento": [{"tipo": t, "modalidade": m,
                        "faturamento": round(v, 2), "m3": round(cruz_m3[(t, m)], 2),
                        "receita_propria": round(cruz_liq[(t, m)], 2)}
                       for (t, m), v in sorted(cruz.items(), key=lambda t: -t[1])],
        "top_clientes": [{
            "nome": sem_documento(c),
            "tipo": tipo_cliente(c),
            "faturamento": round(v["faturamento"], 2),
            "m3": round(v["m3"], 2),
            "preco_m3": round(v["faturamento"] / v["m3"], 2) if v["m3"] else 0,
            "modalidade": "+".join(sorted(v["modalidades"])),
        } for c, v in ranking[:20]],
        "pareto": pareto,
        "recorrencia": dict(faixas),
    }


if __name__ == "__main__":
    caminho = sys.argv[1] if len(sys.argv) > 1 else "relatorio.xls"
    print(json.dumps(analisar(ler(caminho)), ensure_ascii=False, indent=2))
