import xlrd
import openpyxl
from datetime import datetime

ARQUIVO_ENTRADA = "Todos os relatórios.xls"
ARQUIVO_SAIDA = "resumo_ponto.xlsx"
JORNADA_SEMANA = 8 * 60
JORNADA_SABADO = 4 * 60
DIAS_NOMES = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"]

def parse_hora(valor):
    try:
        valor = str(valor).strip()
        h, m = valor.split(":")
        return int(h) * 60 + int(m)
    except:
        return None

def minutos_para_hhmm(minutos):
    sinal = "-" if minutos < 0 else "+"
    minutos = abs(int(minutos))
    return f"{sinal}{minutos // 60:02d}:{minutos % 60:02d}"

def calcular_trabalhado(horarios, dia_semana):
    n = len(horarios)
    jornada = JORNADA_SABADO if dia_semana == 5 else JORNADA_SEMANA

    if n == 0:
        return None, "sem registros — conferir"

    if dia_semana == 5:
        # Sábado: espera entrada e saída apenas
        if n == 1:
            return None, "apenas 1 registro — conferir manual"
        trabalhado = horarios[-1] - horarios[0]
        if trabalhado <= 0 or trabalhado > 480:
            return jornada, "horário suspeito — conferir"
        return trabalhado, None

    if n == 1:
        return None, "apenas 1 registro — conferir"

    if n == 2:
        trabalhado = horarios[-1] - horarios[0] - 120
        if trabalhado <= 0 or trabalhado > 720:
            return jornada, "horário suspeito — conferir"
        return trabalhado, None

    if n == 3:
        # 3 registros: desconta 2h fixas de almoço
        trabalhado = horarios[-1] - horarios[0] - 120
        if trabalhado <= 0 or trabalhado > 720:
            return jornada, "3 registros, horário suspeito — conferir"
        return trabalhado, None

    # 4 ou mais registros
    trabalhado = horarios[-1] - horarios[0]
    for k in range(1, n - 1, 2):
        if k + 1 <= n - 2:
            trabalhado -= (horarios[k + 1] - horarios[k])
    if trabalhado <= 0 or trabalhado > 720:
        return jornada, f"{n} registros, horário suspeito — conferir"
    return trabalhado, None

# Lê arquivo
wb = xlrd.open_workbook(ARQUIVO_ENTRADA)
aba = None
for nome_aba in wb.sheet_names():
    if "log" in nome_aba.lower() or "comparec" in nome_aba.lower():
        aba = wb.sheet_by_name(nome_aba)
        break

print(f"Aba: {aba.name}")

# Extrai mês e ano
mes, ano = None, None
for i in range(5):
    row = aba.row_values(i)
    for val in row:
        val = str(val).strip()
        if "~" in val and "/" in val:
            try:
                data_str = val.split("~")[0].strip()
                data = datetime.strptime(data_str, "%d/%m/%Y")
                mes, ano = data.month, data.year
            except:
                pass

print(f"Período: {mes}/{ano}")

# Mapeamento coluna -> dia
col_para_dia = {}
row_dias = aba.row_values(3)
for j, val in enumerate(row_dias):
    try:
        dia = int(float(str(val).strip()))
        if 1 <= dia <= 31:
            col_para_dia[j] = dia
    except:
        pass

print(f"Dias mapeados: {list(col_para_dia.values())}")

# Processa funcionários
funcionarios = {}

for i in range(aba.nrows):
    row = aba.row_values(i)

    if str(row[0]).strip() == "ID :":
        nome = str(row[9]).strip() if len(row) > 9 else None
        if not nome:
            continue

        linha_horarios = None
        for offset in range(1, 4):
            if i + offset >= aba.nrows:
                break
            row_check = aba.row_values(i + offset)
            conteudo = [str(x).strip() for x in row_check if str(x).strip()]
            if any(":" in c for c in conteudo):
                linha_horarios = row_check
                break

        if linha_horarios is None:
            continue

        if nome not in funcionarios:
            funcionarios[nome] = {
                "total_trabalhado": 0,
                "total_esperado": 0,
                "dias": []
            }

        for col, dia in col_para_dia.items():
            if col >= len(linha_horarios):
                continue

            try:
                dia_semana = datetime(ano, mes, dia).weekday()
            except:
                continue

            if dia_semana == 6:
                continue

            jornada = JORNADA_SABADO if dia_semana == 5 else JORNADA_SEMANA
            data_str = f"{dia:02d}/{mes:02d}/{ano}"
            dia_nome = DIAS_NOMES[dia_semana]

            cell_val = str(linha_horarios[col]).strip()

            # Sem nenhum registro
            if not cell_val:
                funcionarios[nome]["dias"].append({
                    "data": data_str,
                    "dia": dia_nome,
                    "registros": "—",
                    "esperado": "—",
                    "trabalhado": "—",
                    "saldo": "—",
                    "aviso": "sem registros — conferir"
                })
                continue

            horarios_raw = [h.strip() for h in cell_val.split("\n") if h.strip()]
            horarios = [parse_hora(h) for h in horarios_raw]
            horarios = [h for h in horarios if h is not None]

            # Sem registros válidos
            if not horarios:
                funcionarios[nome]["dias"].append({
                    "data": data_str,
                    "dia": dia_nome,
                    "registros": "—",
                    "esperado": "—",
                    "trabalhado": "—",
                    "saldo": "—",
                    "aviso": "sem registros válidos — conferir"
                })
                continue

            registros_str = ", ".join(horarios_raw)

            # 1 registro: não computa o dia
            if len(horarios) == 1:
                sufixo = "conferir manual" if dia_semana == 5 else "conferir"
                funcionarios[nome]["dias"].append({
                    "data": data_str,
                    "dia": dia_nome,
                    "registros": registros_str,
                    "esperado": "—",
                    "trabalhado": "—",
                    "saldo": "—",
                    "aviso": f"apenas 1 registro — {sufixo}"
                })
                continue

            # Dia com registros suficientes
            funcionarios[nome]["total_esperado"] += jornada

            trabalhado, aviso = calcular_trabalhado(horarios, dia_semana)

            if trabalhado is not None:
                funcionarios[nome]["total_trabalhado"] += trabalhado

            trab_str = minutos_para_hhmm(trabalhado) if trabalhado is not None else "—"
            saldo_str = minutos_para_hhmm(trabalhado - jornada) if trabalhado is not None else "—"

            funcionarios[nome]["dias"].append({
                "data": data_str,
                "dia": dia_nome,
                "registros": registros_str,
                "esperado": minutos_para_hhmm(jornada),
                "trabalhado": trab_str,
                "saldo": saldo_str,
                "aviso": aviso or ""
            })

# Gera Excel
wb_out = openpyxl.Workbook()
ws = wb_out.active
ws.title = "Resumo Ponto"
ws.append(["Nome", "Total Trabalhado", "Total Esperado", "Saldo", "Qtd Avisos"])

print("\nResumo:")
for nome, f in sorted(funcionarios.items()):
    saldo = f["total_trabalhado"] - f["total_esperado"]
    qtd_avisos = sum(1 for d in f["dias"] if d["aviso"])
    ws.append([
        nome,
        minutos_para_hhmm(f["total_trabalhado"]),
        minutos_para_hhmm(f["total_esperado"]),
        minutos_para_hhmm(saldo),
        qtd_avisos
    ])
    print(f"  {nome}: saldo {minutos_para_hhmm(saldo)} | avisos: {qtd_avisos}")

    # Aba detalhada para cada funcionário
    ws_func = wb_out.create_sheet(title=nome[:31])
    ws_func.append(["Data", "Dia", "Registros", "Esperado", "Trabalhado", "Saldo Dia", "Observação"])
    for d in f["dias"]:
        ws_func.append([
            d["data"],
            d["dia"],
            d["registros"],
            d["esperado"],
            d["trabalhado"],
            d["saldo"],
            d["aviso"]
        ])

wb_out.save(ARQUIVO_SAIDA)
print(f"\n✓ Salvo: {ARQUIVO_SAIDA}")
