import xlrd
import openpyxl
from datetime import datetime

ARQUIVO_ENTRADA = "Todos os relatórios.xls"
ARQUIVO_SAIDA = "resumo_ponto.xlsx"
JORNADA_SEMANA = 8 * 60   # minutos
JORNADA_SABADO = 4 * 60   # minutos

def parse_hora(valor):
    """Converte HH:MM para minutos."""
    try:
        valor = str(valor).strip()
        h, m = valor.split(":")
        return int(h) * 60 + int(m)
    except:
        return None

def minutos_para_hhmm(minutos):
    """Converte minutos para +HH:MM ou -HH:MM."""
    sinal = "-" if minutos < 0 else "+"
    minutos = abs(int(minutos))
    return f"{sinal}{minutos // 60:02d}:{minutos % 60:02d}"

def calcular_trabalhado(horarios, dia_semana):
    """
    Calcula minutos trabalhados a partir de uma lista de horários.
    Retorna (minutos, aviso)
    """
    n = len(horarios)

    if n == 0:
        return None, None  # sem registro, ignora

    if n == 1:
        # Só um registro — incompleto
        jornada = JORNADA_SABADO if dia_semana == 5 else JORNADA_SEMANA
        return jornada, f"apenas 1 registro — conferir"

    if dia_semana == 5:
        # Sábado — entrada e saída, sem almoço
        trabalhado = horarios[-1] - horarios[0]
        if trabalhado <= 0 or trabalhado > 480:
            jornada = JORNADA_SABADO
            return jornada, f"horário suspeito — conferir"
        return trabalhado, None

    # Dia de semana
    if n == 2:
        # Entrada e saída sem almoço — desconta 2h
        trabalhado = horarios[-1] - horarios[0] - 120
        if trabalhado <= 0 or trabalhado > 720:
            return JORNADA_SEMANA, f"horário suspeito — conferir"
        return trabalhado, None

    if n >= 3:
        # Primeiro = entrada, último = saída
        trabalhado = horarios[-1] - horarios[0]
        # Desconta intervalos do meio (saída almoço até retorno)
        for k in range(1, n - 1, 2):
            if k + 1 <= n - 2:
                trabalhado -= (horarios[k + 1] - horarios[k])
        if trabalhado <= 0 or trabalhado > 720:
            return JORNADA_SEMANA, f"{n} registros, horário suspeito — conferir"
        return trabalhado, None

# Lê arquivo
wb = xlrd.open_workbook(ARQUIVO_ENTRADA)

# Encontra aba
aba = None
for nome_aba in wb.sheet_names():
    if "log" in nome_aba.lower() or "comparec" in nome_aba.lower():
        aba = wb.sheet_by_name(nome_aba)
        break

print(f"Aba: {aba.name}")

# Extrai mês e ano do cabeçalho
mes, ano = None, None
for i in range(5):
    for j in range(10):
        val = str(aba.cell_value(i, j))
        if "~" in val and "/" in val:
            try:
                data_str = val.split("~")[0].strip()
                data = datetime.strptime(data_str, "%d/%m/%Y")
                mes, ano = data.month, data.year
            except:
                pass

print(f"Período: {mes}/{ano}")

# Descobre quais colunas correspondem a quais dias
# Usa a primeira linha de números (linha 3, índice 3)
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
    row_str = " ".join([str(x) for x in row])

    # Detecta linha de cabeçalho do funcionário
    if "Nome" in row_str and "ID" in row_str and "Dept" in row_str:
        # Extrai nome
        nome = None
        for j in range(len(row)):
            if str(row[j]).strip() == "Nome":
                for k in range(j + 1, min(j + 5, len(row))):
                    val = str(row[k]).strip()
                    if val and val != ":" and val != "Nome":
                        nome = val
                        break
                break

        if not nome:
            continue

        # Próxima linha com dados é a de horários
        if i + 1 >= aba.nrows:
            continue

        # Pula linhas de números (1-20) até achar horários
        linha_horarios = None
        for offset in range(1, 4):
            if i + offset >= aba.nrows:
                break
            row_check = aba.row_values(i + offset)
            conteudo = [str(x).strip() for x in row_check if str(x).strip()]
            # Se tem ":" é linha de horários
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

            dia_semana = datetime(ano, mes, dia).weekday()

            # Ignora domingo
            if dia_semana == 6:
                continue

            jornada = JORNADA_SABADO if dia_semana == 5 else JORNADA_SEMANA
            funcionarios[nome]["total_esperado"] += jornada

            # Extrai horários da célula (separados por \n)
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