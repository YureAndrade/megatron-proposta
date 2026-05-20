#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador de cotacao - ORCAMENTO MATERIAIS ELETRICOS (CABOS)
Cliente: JUBA ATACADO - PONTES E LACERDA / CNPJ 03.550.647/0008-51

Le a tabela de precos (ALUMINIO - COBRE.xlsx, aba ENERGIA) e resolve, para
cada item do orcamento, o CODTOTVS e a quantidade prontos para lancamento.

Padrao adotado: codigo POR METRO (sufixo F0001) com a quantidade em metros
exatamente como pedida no orcamento. Onde nao existe codigo por metro
(ex.: fio 750V 1,5mm), usa-se rolo de 100m (sufixo R0100).
Itens sem equivalente na tabela ficam marcados como 'A DEFINIR'.
"""
import csv
import os
import openpyxl

CATALOGO = os.path.join(os.path.dirname(__file__), "ALUMÍNIO - COBRE.xlsx")
SAIDA = os.path.join(os.path.dirname(__file__), "cotacao_orcamento_eletrica.csv")


def carregar_catalogo():
    """Retorna dict {CODTOTVS: (descricao, preco)} a partir da aba ENERGIA."""
    wb = openpyxl.load_workbook(CATALOGO, data_only=True)
    cat = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            cod_totvs = None
            desc = None
            preco = None
            for c in row:
                if isinstance(c, str) and len(c) >= 10 and any(ch.isdigit() for ch in c) \
                        and not any(k in c.upper() for k in ["CABO", "BOB", "CORDAO", "CORDÃO", "FIO"]):
                    cod_totvs = cod_totvs or c.strip()
                if isinstance(c, str) and any(k in c.upper() for k in ["CABO", "BOB", "CORDAO", "CORDÃO", "FIO"]):
                    desc = desc or c.strip()
            floats = [c for c in row if isinstance(c, float)]
            if floats:
                preco = floats[-1]
            if cod_totvs and desc:
                cat[cod_totvs] = (desc, preco)
    return cat


# CODTOTVS HEPR unipolar flexivel 0,6/1kV: 0111<NN>F0001<COR>00
GAUGE_SEG = {1.5: "01", 2.5: "02", 4: "03", 6: "04", 10: "05", 16: "06",
             25: "07", 35: "08", 50: "09", 70: "10", 95: "11", 120: "12",
             150: "13", 185: "14", 240: "15", 300: "16"}


def uni_hepr(bitola, cor):
    """Codigo POR METRO do cabo flexivel HEPR unipolar (cor: PT/AZ/VD)."""
    return f"0111{GAUGE_SEG[bitola]}F0001{cor}00"


def linha(secao, descricao, cor, bitola, qtd, unid, codigo, obs=""):
    return dict(secao=secao, descricao=descricao, cor=cor, bitola=bitola,
                qtd=qtd, unid=unid, codigo=codigo, obs=obs)


def montar_orcamento():
    itens = []

    # ---- A) CABO TRIPOLAR (cobre) - PP 3 VIAS HEPR (so existe em PT) ----
    itens.append(linha("Cabo Tripolar (cobre)", "Cabo PP 3 vias HEPR 0,6/1kV 3x10mm", "PRETO", "10mm", 1121.87, "m", "011127F0001PT00",
                        "Tabela so tem tripolar HEPR na cor preta"))
    itens.append(linha("Cabo Tripolar (cobre)", "Cabo PP 3 vias HEPR 0,6/1kV 3x2,5mm", "PRETO", "2.5mm", 600.42, "m", "011124F0001PT00",
                        "Tabela so tem tripolar HEPR na cor preta"))

    # ---- B) SISTEMA ELETRICO - BOMBEIRO ----
    itens.append(linha("Bombeiro", "Cabo 3 vias 1,5mm BLINDADO (shield) vermelho", "VERMELHO", "1.5mm", 2500, "m", "A DEFINIR",
                        "NAO HA cabo blindado na tabela. Equivalente s/ blindagem: PP 3 vias 1,5mm 011123F0001PT00"))
    for cor, cod in [("VERMELHO", "010404R0100VM00"), ("PRETO", "010404R0100PT00"), ("VERDE", "010404R0100VD00")]:
        itens.append(linha("Bombeiro", "Fio flexivel 750V 1,5mm", cor, "1.5mm", 25, "rolo 100m", cod,
                            "750V 1,5mm so vendido em rolo 100m (2.500m = 25 rolos)"))

    # ---- C) CABO UNIPOLAR HEPR 0,6/1kV (por metro) ----
    unipolar = {
        2.5: {"AM": 500, "AZ": 6123.6, "PT": 6307.50, "VD": 2255.3},
        4:   {"AM": 300, "AZ": 5368.7, "PT": 5516.1, "VD": 2474.7},
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
    cor_nome = {"PT": "PRETO", "AZ": "AZUL", "VD": "VERDE", "AM": "AMARELO"}
    for g in sorted(unipolar):
        for cor in ["AM", "AZ", "PT", "VD"]:
            if cor not in unipolar[g]:
                continue
            q = unipolar[g][cor]
            if cor == "AM":  # amarelo nao existe em HEPR
                itens.append(linha("Cabo Unipolar", f"Cabo flexivel HEPR 0,6/1kV {g}mm", "AMARELO",
                                   f"{g}mm", q, "m", "A DEFINIR",
                                   "HEPR nao tem AMARELO na tabela (so PT/AZ/VD)"))
            else:
                itens.append(linha("Cabo Unipolar", f"Cabo flexivel HEPR 0,6/1kV {g}mm",
                                   cor_nome[cor], f"{g}mm", q, "m", uni_hepr(g, cor)))

    # ---- D) MAT ATERRAMENTO / ENTRADA / MUFLA / BT ----
    itens.append(linha("Aterramento/Entrada", "Cabo XLPE 25mm cobre", "-", "25mm", 25, "m", "A DEFINIR",
                        "Tabela nao tem XLPE de COBRE. Equivalente: HEPR 25mm 011107F0001PT00"))
    itens.append(linha("Aterramento/Entrada", "Cabo cobre flexivel HEPR 10mm 0,6/1kV", "VERDE", "10mm", 15, "m", "011105F0001VD00"))
    itens.append(linha("Aterramento/Entrada", "Cabo cobre flexivel HEPR 50mm 0,6/1kV", "VERDE", "50mm", 10, "m", "011109F0001VD00"))
    itens.append(linha("Aterramento/Entrada", "Cabo cobre flexivel HEPR 95mm 0,6/1kV", "AZUL", "95mm", 252, "m", "011111F0001AZ00"))
    itens.append(linha("Aterramento/Entrada", "Cabo cobre flexivel HEPR 185mm 0,6/1kV", "PRETO", "185mm", 760, "m", "011114F0001PT00"))
    itens.append(linha("Aterramento/Entrada", "Cabo Mufla Media Tensao 35mm 15/20kV", "-", "35mm", 120, "m", "A DEFINIR",
                        "Tabela nao tem cabo de media tensao/mufla. Sob consulta"))
    return itens


def main():
    cat = carregar_catalogo()
    itens = montar_orcamento()
    campos = ["secao", "descricao", "cor", "bitola", "qtd", "unid", "codigo",
              "descricao_totvs", "preco_unit", "obs"]
    total = 0.0
    with open(SAIDA, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=campos, delimiter=";")
        w.writeheader()
        for it in itens:
            desc_totvs, preco = cat.get(it["codigo"], ("", None))
            it["descricao_totvs"] = desc_totvs
            it["preco_unit"] = f"{preco:.4f}".replace(".", ",") if preco else ""
            if preco and it["codigo"] != "A DEFINIR":
                total += preco * it["qtd"]
            w.writerow(it)
    print(f"Gerado: {SAIDA}  ({len(itens)} itens)")
    print(f"Valor de referencia (tabela, itens resolvidos): R$ {total:,.2f}")


if __name__ == "__main__":
    main()
