import xlrd
import openpyxl
from datetime import datetime

ARQUIVO_ENTRADA = "Todos os relatórios.xls"
ARQUIVO_SAIDA = "resumo_ponto.xlsx"
JORNADA_SEMANA = 8 * 60
JORNADA_SABADO = 4 * 60

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
        return None, None

    if n == 1:
        return jornada, "apenas 1 registro — conferir"

    if dia_semana == 5:
        trabalhado = horarios[-1] - horarios[0]
        if trabalhado <= 0 or trabalhado > 480:
            return jornada, "horário suspeito — conferir"
        return trabalhado, None

    if n == 2:
        trabalhado = horarios[-1] - horarios[0] - 120
        if trabalhado <= 0 or trabalhado > 720:
            return jornada, "horário suspeito — conferir"
        return trabalhado, None

    # 3 ou mais registros
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

# Mapeamento coluna -> dia (usa primeira linha de números, índice 3)
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

    # Detecta linha de cabeçalho: tem "ID :" na coluna 0
    if str(row[0]).strip() == "ID :":
        # Nome está na coluna 9
        nome = str(row[9]).strip() if len(row) > 9 else None

        if not nome:
            continue

        # Próxima linha com horários
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
                "avisos": []
            }

        # Processa cada dia
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
            funcionarios[nome]["total_esperado"] += jornada

            cell_val = str(linha_horarios[col]).strip()
            if not cell_val:
                continue

            horarios_raw = [h.strip() for h in cell_val.split("\n") if h.strip()]
            horarios = [parse_hora(h) for h in horarios_raw]
            horarios = [h for h in horarios if h is not None]

            if not horarios:
                continue

            trabalhado, aviso = calcular_trabalhado(horarios, dia_semana)

            if trabalhado is not None:
                funcionarios[nome]["total_trabalhado"] += trabalhado
            if aviso:
                data_str = f"{dia:02d}/{mes:02d}/{ano}"
                funcionarios[nome]["avisos"].append(f"{data_str}: {aviso}")

# Gera Excel
wb_out = openpyxl.Workbook()
ws = wb_out.active
ws.title = "Resumo Ponto"
ws.append(["Nome", "Total Trabalhado", "Total Esperado", "Saldo", "Avisos"])

print("\nResumo:")
for nome, f in sorted(funcionarios.items()):
    saldo = f["total_trabalhado"] - f["total_esperado"]
    avisos = " | ".join(f["avisos"]) if f["avisos"] else ""
    ws.append([
        nome,
        minutos_para_hhmm(f["total_trabalhado"]),
        minutos_para_hhmm(f["total_esperado"]),
        minutos_para_hhmm(saldo),
        avisos
    ])
    print(f"  {nome}: saldo {minutos_para_hhmm(saldo)} | avisos: {len(f['avisos'])}")

wb_out.save(ARQUIVO_SAIDA)
print(f"\n✓ Salvo: {ARQUIVO_SAIDA}")