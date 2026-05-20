#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador de cotacao - ORCAMENTO MATERIAIS ELETRICOS (CABOS)
Cliente: JUBA ATACADO - PONTES E LACERDA / CNPJ 03.550.647/0008-51

Le a tabela de precos (ALUMINIO - COBRE.xlsx) e resolve, para cada item do
orcamento, o(s) CODTOTVS e a quantidade prontos para lancamento no TOTVS.

Regras definidas pelo cliente:
  * Unidade: ROLO/BOBINA FECHADA, arredondando a metragem pra cima e
    minimizando a sobra. Onde nao existe embalagem fechada na tabela
    (bitolas >= 50mm, alguns tripolares), usa-se o codigo POR METRO (F0001)
    com a metragem exata.
  * Amarelo: HEPR nao tem amarelo -> substituido por PRETO (mesma metragem).
  * Cabo XLPE 25mm cobre: substituido por HEPR flexivel 25mm cobre.
  * Cabo 3 vias blindado e mufla/MT: nao existem na tabela -> 'SOB CONSULTA'.
"""
import csv
import math
import os
import openpyxl

CATALOGO = os.path.join(os.path.dirname(__file__), "ALUMÍNIO - COBRE.xlsx")
SAIDA = os.path.join(os.path.dirname(__file__), "cotacao_orcamento_eletrica.csv")


def carregar_catalogo():
    """Indexa a tabela. Retorna:
       info[cod] = (descricao, preco)
       pack[(prefixo6, cor)] = lista de (comprimento_m, cod, tipo)  tipo in F/R/B
    """
    wb = openpyxl.load_workbook(CATALOGO, data_only=True)
    info, pack = {}, {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            cod = desc = preco = None
            for c in row:
                if isinstance(c, str) and len(c) == 15 and c[6] in "FRB" \
                        and c[:6].isdigit() and c[7:11].isdigit():
                    cod = cod or c.strip()
                if isinstance(c, str) and any(k in c.upper() for k in
                                              ["CABO", "BOB", "CORDAO", "CORDÃO", "FIO"]):
                    desc = desc or c.strip()
            floats = [c for c in row if isinstance(c, float)]
            if floats:
                preco = floats[-1]
            if cod and desc:
                info[cod] = (desc, preco)
                comp = int(cod[7:11])
                cor = cod[11:13]
                pack.setdefault((cod[:6], cor), []).append((comp, cod, cod[6]))
    return info, pack


def empacotar(need_m, denoms):
    """denoms: lista de (comprimento_m, cod) fechados (R/B), multiplos de 100.
    Retorna (total_m, [(cod, qtd), ...]) cobrindo need_m com a menor sobra e
    o menor numero de pecas. Retorna None se nao houver embalagem fechada."""
    if not denoms:
        return None
    code_by_u = {}
    for comp, cod in sorted(denoms):
        code_by_u.setdefault(comp // 100, cod)   # 1 codigo por comprimento
    units = sorted(code_by_u)
    T = math.ceil(round(need_m, 3) / 100)
    hi = T + max(units)
    INF = float("inf")
    dp = [INF] * (hi + 1)
    pick = [None] * (hi + 1)
    dp[0] = 0
    for a in range(1, hi + 1):
        for u in units:
            if a - u >= 0 and dp[a - u] + 1 < dp[a]:
                dp[a] = dp[a - u] + 1
                pick[a] = u
    best = next((a for a in range(T, hi + 1) if dp[a] < INF), None)
    counts = {}
    a = best
    while a > 0:
        counts[pick[a]] = counts.get(pick[a], 0) + 1
        a -= pick[a]
    itens = [(code_by_u[u], q) for u, q in sorted(counts.items(), reverse=True)]
    return best * 100, itens


GAUGE_SEG = {1.5: "01", 2.5: "02", 4: "03", 6: "04", 10: "05", 16: "06",
             25: "07", 35: "08", 50: "09", 70: "10", 95: "11", 120: "12",
             150: "13", 185: "14", 240: "15", 300: "16"}
COR_NOME = {"PT": "PRETO", "AZ": "AZUL", "VD": "VERDE", "VM": "VERMELHO"}


def main():
    info, pack = carregar_catalogo()
    linhas = []   # cada linha = dict pronto pro CSV

    def resolver(secao, desc, bitola, cor, need_m, prefixo, obs=""):
        """Gera as linhas (embalagem fechada ou por metro) para um requisito."""
        chave = (prefixo, cor)
        fechados = [(c, cod) for c, cod, t in pack.get(chave, []) if t in "RB"]
        fcode = next((cod for c, cod, t in pack.get(chave, []) if t == "F"), None)
        emb = empacotar(need_m, fechados)
        if emb:
            total, itens = emb
            sobra = total - need_m
            for cod, q in itens:
                d, p = info[cod]
                comp = int(cod[7:11])
                linhas.append(dict(secao=secao, descricao=desc, bitola=bitola,
                    cor=COR_NOME.get(cor, cor), codigo=cod, qtd=q,
                    unid=("rolo 100m" if cod[6] == "R" else f"bobina {comp}m"),
                    metros=q * comp, preco_unit=fmt(p),
                    desc_totvs=d, obs=obs))
            linhas[-1]["obs"] = (obs + f" | pedido {need_m:g}m -> {total}m (sobra {sobra:g}m)").strip(" |")
        else:  # so por metro
            d, p = info[fcode]
            linhas.append(dict(secao=secao, descricao=desc, bitola=bitola,
                cor=COR_NOME.get(cor, cor), codigo=fcode, qtd=need_m, unid="m",
                metros=need_m, preco_unit=fmt(p), desc_totvs=d,
                obs=(obs + " | so existe por metro na tabela").strip(" |")))

    def pendente(secao, desc, bitola, qtd, obs):
        linhas.append(dict(secao=secao, descricao=desc, bitola=bitola, cor="-",
            codigo="SOB CONSULTA", qtd=qtd, unid="m", metros=qtd,
            preco_unit="", desc_totvs="", obs=obs))

    # ---- A) Cabo Tripolar (cobre) - PP 3 vias HEPR (so PT) ----
    resolver("Cabo Tripolar", "PP 3 vias HEPR 0,6/1kV 3x10mm", "10mm", "PT", 1121.87, "011127",
             "tripolar HEPR so existe em preto")
    resolver("Cabo Tripolar", "PP 3 vias HEPR 0,6/1kV 3x2,5mm", "2.5mm", "PT", 600.42, "011124",
             "tripolar HEPR so existe em preto")

    # ---- B) Bombeiro ----
    pendente("Bombeiro", "Cabo 3 vias 1,5mm BLINDADO (shield) vermelho", "1.5mm", 2500,
             "Nao ha cabo blindado na tabela - cotar com fornecedor especifico")
    for cor in ["VM", "PT", "VD"]:
        resolver("Bombeiro", "Fio flexivel 750V 1,5mm", "1.5mm", cor, 2500, "010404",
                 "fio bombeiro 750V")

    # ---- C) Cabo Unipolar HEPR 0,6/1kV ----
    unipolar = {
        2.5: {"PT": 6307.50 + 500, "AZ": 6123.6, "VD": 2255.3},   # +500 amarelo->preto
        4:   {"PT": 5516.1 + 300, "AZ": 5368.7, "VD": 2474.7},    # +300 amarelo->preto
        6:   {"AZ": 400, "PT": 400, "VD": 400},
        10:  {"AZ": 250, "PT": 500, "VD": 250},
        16:  {"AZ": 200, "PT": 500, "VD": 500},
        25:  {"AZ": 100, "PT": 300, "VD": 100},
        35:  {"AZ": 20, "PT": 50, "VD": 40},
        50:  {"AZ": 200, "PT": 450, "VD": 300},
        70:  {"AZ": 30, "PT": 100, "VD": 400},
        95:  {"AZ": 300, "PT": 850, "VD": 60},
        120: {"AZ": 400, "PT": 1200},
    }
    for g in sorted(unipolar):
        for cor in ["PT", "AZ", "VD"]:
            if cor not in unipolar[g]:
                continue
            obs = "inclui amarelo->preto" if (cor == "PT" and g in (2.5, 4)) else ""
            resolver("Cabo Unipolar", f"Cabo flexivel HEPR 0,6/1kV {g}mm", f"{g}mm",
                     cor, unipolar[g][cor], "0111" + GAUGE_SEG[g], obs)

    # ---- D) Aterramento / Entrada / Mufla ----
    resolver("Aterramento", "Cabo HEPR 25mm cobre (subst. XLPE)", "25mm", "PT", 25, "011107",
             "substitui 'XLPE 25mm cobre' (tabela nao tem XLPE de cobre)")
    resolver("Aterramento", "Cabo cobre flexivel HEPR 10mm", "10mm", "VD", 15, "011105")
    resolver("Aterramento", "Cabo cobre flexivel HEPR 50mm", "50mm", "VD", 10, "011109")
    resolver("Aterramento", "Cabo cobre flexivel HEPR 95mm", "95mm", "AZ", 252, "011111")
    resolver("Aterramento", "Cabo cobre flexivel HEPR 185mm", "185mm", "PT", 760, "011114")
    pendente("Aterramento", "Cabo Mufla Media Tensao 35mm 15/20kV", "35mm", 120,
             "Nao ha cabo de media tensao/mufla na tabela - cotar a parte")

    # ---- saida ----
    campos = ["secao", "descricao", "bitola", "cor", "codigo", "qtd", "unid",
              "metros", "preco_unit", "total_linha", "desc_totvs", "obs"]
    total = 0.0
    with open(SAIDA, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=campos, delimiter=";")
        w.writeheader()
        for ln in linhas:
            d, p = info.get(ln["codigo"], ("", None))
            tl = p * ln["qtd"] if p else 0.0   # preco e por embalagem (ou por metro)
            ln["total_linha"] = fmt(round(tl, 2)) if p else ""
            total += tl
            w.writerow(ln)
    print(f"Gerado: {SAIDA}  ({len(linhas)} linhas)")
    print(f"Valor de referencia (tabela, itens resolvidos, s/ desconto): R$ {total:,.2f}")


def fmt(p):
    return f"{p:.4f}".replace(".", ",") if p else ""


if __name__ == "__main__":
    main()
