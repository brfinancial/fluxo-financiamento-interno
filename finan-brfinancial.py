import streamlit as st
from pathlib import Path
from io import BytesIO
import calendar
from datetime import datetime as dt, time
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# --- Auxiliares de taxa externa ---
def load_taxas(filepath: str) -> dict:
    taxas = {}
    path = Path(filepath)
    if not path.exists():
        st.error(f"Arquivo de taxas não encontrado: {filepath}")
        return taxas
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read().strip()
    blocos = [b.strip() for b in content.split("\n\n") if b.strip()]
    for bloco in blocos:
        linhas = bloco.splitlines()
        nome = linhas[0].strip()
        taxas[nome] = {}
        for linha in linhas[1:]:
            if '=' in linha:
                chave, valor = linha.split('=', 1)
                try:
                    taxas[nome][chave.strip()] = float(valor.strip())
                except ValueError:
                    taxas[nome][chave.strip()] = valor.strip()
    return taxas

# --- Formatação Excel ---
HEADER_FILL = PatternFill(start_color="FFD3D3D3", end_color="FFD3D3D3", fill_type="solid")
DATE_FORMAT = 'dd/mm/yyyy'
CURRENCY_FORMAT = '"R$" #,##0.00'

def adjust_day(date, preferred_day):
    try:
        return date.replace(day=preferred_day)
    except ValueError:
        last = calendar.monthrange(date.year, date.month)[1]
        return date.replace(day=last)

class PaymentTracker:
    def __init__(self, dia_pagamento, taxa_juros):
        self.last_date = None
        self.dia = dia_pagamento
        self.taxa = taxa_juros
    def calculate(self, current_date, saldo):
        if self.last_date is None:
            self.last_date = current_date
            return 0.0, 0, 0.0
        dias = (current_date - self.last_date).days
        taxa_eff = self.taxa * (dias / 30)
        juros = saldo * taxa_eff
        self.last_date = current_date
        return juros, dias, taxa_eff

# --- App Streamlit ---
def main():
    st.set_page_config(page_title="Gerador de Planilha de Financiamento", layout="centered")
    st.title("Gerador de Financiamento — Br Financial")

    # Carrega taxas
    taxas_por_emp = load_taxas('taxas.txt')

    # Entradas
    cliente = st.text_input("Cliente")
    valor_imovel = st.number_input("Valor do imóvel (R$)", min_value=0.0, step=0.01, format="%.2f")
    dia_pag = st.number_input("Dia p/ parcelas mensais (1-31)", 1, 31, 1)

    empreendimento = st.selectbox("Empreendimento", list(taxas_por_emp.keys()))
    sel = taxas_por_emp.get(empreendimento, {})
    TAXA_INCC = sel.get('TAXA_INCC', 0.0)
    TAXA_IPCA = sel.get('TAXA_IPCA', 0.0)
    taxa_pre = sel.get('taxa_pre', 0.0)
    taxa_pos = sel.get('taxa_pos', 0.0)
    taxas_extras = [
        {'pct': v, 'periodo': 'pré-entrega'} if 'INCC' in k else {'pct': v, 'periodo': 'pós-entrega'}
        for k, v in sel.items() if k.endswith('_PCT') and k != 'TAXA_SEGURO_PRESTAMISTA_PCT'
    ]

    data_base = dt.combine(st.date_input("Data base (assinatura)"), time())
    capacidade_pre = st.number_input("Capacidade pré-entrega (R$)", 0.0, step=0.01)
    data_inicio_pre = dt.combine(st.date_input("Início pré-entrega"), time())
    data_entrega = dt.combine(st.date_input("Data de entrega"), time())
    fgts = st.number_input("FGTS (abatimento)", 0.0, step=0.01)
    fin_banco = st.number_input("Financiamento banco (abatimento)", 0.0, step=0.01)
    capacidade_pos = st.number_input("Capacidade pós-entrega (R$)", 0.0, step=0.01) - st.number_input("Parcela banco (R$)", 0.0, step=0.01)

    # Não recorrentes
    st.subheader("Pagamentos adicionais")
    non_rec = []
    n_nr = st.number_input("Nº de adicionais", 0, step=1)
    for i in range(n_nr):
        d = dt.combine(st.date_input(f"Data add {i+1}", key=f"nr_d{i}"), time())
        v = st.number_input(f"Valor add {i+1}", 0.0, step=0.01, key=f"nr_v{i}")
        desc = st.text_input(f"Desc add {i+1}", key=f"nr_desc{i}")
        assoc = st.checkbox("Associar à mensal?", key=f"nr_assoc{i}")
        if assoc: d = adjust_day(d, dia_pag)
        non_rec.append({'data': d, 'tipo': desc, 'valor': v, 'assoc': assoc})

    # Séries semestrais
    st.subheader("Séries semestrais")
    semi_series = []
    for i in range(st.number_input("Nº semestrais", 0, step=1)):
        d0 = dt.combine(st.date_input(f"Data semi {i+1}", key=f"s_d{i}"), time())
        v = st.number_input(f"Valor semi {i+1}", 0.0, step=0.01, key=f"s_v{i}")
        assoc = st.checkbox("Assoc. à mensal?", key=f"s_a{i}")
        semi_series.append({'d0': d0, 'v': v, 'assoc': assoc})

    # Séries anuais
    st.subheader("Séries anuais")
    annual_series = []
    for i in range(st.number_input("Nº anuais", 0, step=1)):
        d0 = dt.combine(st.date_input(f"Data anual {i+1}", key=f"a_d{i}"), time())
        v = st.number_input(f"Valor anual {i+1}", 0.0, step=0.01, key=f"a_v{i}")
        assoc = st.checkbox("Assoc. à mensal?", key=f"a_a{i}")
        annual_series.append({'d0': d0, 'v': v, 'assoc': assoc})

    if st.button("Gerar Planilha"):
        # Monta listas próprias
        sem_rec = []
        for s in semi_series:
            for n in range(100):
                d = s['d0'] + relativedelta(months=6*n)
                if s['assoc']: d = adjust_day(d, dia_pag)
                sem_rec.append({'data': d, 'tipo': 'Pagamento Semestral', 'valor': s['v'], 'assoc': s['assoc']})
        ann_rec = []
        for a in annual_series:
            for n in range(100):
                d = a['d0'] + relativedelta(years=n)
                if a['assoc']: d = adjust_day(d, dia_pag)
                ann_rec.append({'data': d, 'tipo': 'Pagamento Anual', 'valor': a['v'], 'assoc': a['assoc']})

        # Extendendo non_rec com não-associados
        non_rec += [e for e in sem_rec + ann_rec if not e['assoc']]

        # Separa pré / pós
        pre_nr = sorted([e for e in non_rec if e['data'] < data_entrega], key=lambda x: x['data'])
        post_nr = sorted([e for e in non_rec if e['data'] >= data_entrega], key=lambda x: x['data'])

        eventos = []
        saldo = valor_imovel

        # Data base
        eventos.append({
            'data': data_base, 'parcela': '', 'tipo': 'Data-Base',
            'valor': 0.0, 'juros': 0.0, 'dias_corridos': 0, 'taxa_efetiva': 0.0,
            'incc': 0.0, 'ipca': 0.0, 'taxas_extra': [0.0]*len(taxas_extras),
            'Total de mudança (R$)': 0.0, 'saldo': saldo
        })

        # PRÉ-ENTREGA
        tp = PaymentTracker(dia_pag, taxa_pre); tp.last_date = data_base
        mc_pre = sc_pre = ac_pre = 0
        prev = cursor = data_inicio_pre

        while True:
            d_evt = adjust_day(cursor, dia_pag)
            if d_evt >= data_entrega: break

            # não-associados
            for ev in [e for e in pre_nr if not e['assoc'] and prev < e['data'] <= d_evt]:
                if ev['tipo']=='Pagamento Semestral':
                    sc_pre+=1; label=f"{sc_pre}ª Parcela Semestral Pré-Entrega"
                    juros=dias=taxa_eff=0.0; saldo-=ev['valor']
                elif ev['tipo']=='Pagamento Anual':
                    ac_pre+=1; label=f"{ac_pre}ª Parcela Anual Pré-Entrega"
                    juros=dias=taxa_eff=0.0; saldo-=ev['valor']
                elif ev['tipo'].startswith("Parcela Mensal"):
                    mc_pre+=1; label=f"{mc_pre}ª Parcela Mensal Pré-Entrega"
                    juros,dias,taxa_eff = tp.calculate(ev['data'], saldo)
                    incc = saldo * TAXA_INCC
                    extras = sum(saldo*t['pct'] for t in taxas_extras if t['periodo']=='pré-entrega')
                    abat = ev['valor'] - juros - incc - extras; saldo-=abat
                    eventos.append({**ev, 'tipo':label, 'juros':juros, 'dias_corridos':dias,
                                'taxa_efetiva':taxa_eff, 'incc':incc if 'incc' in locals() else 0.0,
                                'ipca':0.0, 'taxas_extra':[0.0]*len(taxas_extras),
                                'Total de mudança (R$)':ev['valor'], 'saldo':saldo})
                else:
                    label = ev['tipo']
                    juros, dias, taxa_eff = tp.calculate(ev['data'], saldo)
                    incc = saldo * TAXA_INCC
                    extras = sum(saldo * t['pct'] for t in taxas_extras if t['periodo']=='pré-entrega')
                    abat = ev['valor'] - juros - incc - extras
                    saldo -= abat
                
                    eventos.append({
                        'data': ev['data'],
                        'parcela': '',
                        'tipo': ev['tipo'],
                        'valor': ev['valor'],
                        'juros': juros,
                        'dias_corridos': dias,
                        'taxa_efetiva': taxa_eff,
                        'incc': incc if 'incc' in locals() else 0.0,
                        'ipca': 0.0,
                        'taxas_extra': [0.0] * len(taxas_extras),
                        'Total de mudança (R$)': abat if 'abat' in locals() else ev['valor'],
                        'saldo': saldo
                    })
                    
            # associados semestrais
            for ev in [e for e in sem_rec if e['assoc'] and prev < e['data'] <= d_evt]:
                sc_pre+=1; label=f"{sc_pre}ª Parcela Semestral Pré-Entrega"
                saldo-=ev['valor']
                eventos.append({**ev,'tipo':label,'juros':0.0,'dias_corridos':0,'taxa_efetiva':0.0,
                                'incc':0.0,'ipca':0.0,'taxas_extra':[0.0]*len(taxas_extras),
                                'Total de mudança (R$)':ev['valor'],'saldo':saldo})
            # associados anuais
            for ev in [e for e in ann_rec if e['assoc'] and prev < e['data'] <= d_evt]:
                ac_pre+=1; label=f"{ac_pre}ª Parcela Anual Pré-Entrega"
                saldo-=ev['valor']
                eventos.append({**ev,'tipo':label,'juros':0.0,'dias_corridos':0,'taxa_efetiva':0.0,
                                'incc':0.0,'ipca':0.0,'taxas_extra':[0.0]*len(taxas_extras),
                                'Total de mudança (R$)':ev['valor'],'saldo':saldo})
            
            # — Eventos não-recorrentes ASSOCIADOS pré‑entrega —
            for ev in [e for e in pre_nr if e['assoc'] and prev < e['data'] <= d_evt]:
                # data já ajustada em pre_nr
                eventos.append({
                    'data': ev['data'],
                    'parcela': '',
                    'tipo': ev['tipo'],            # mantém o texto do usuário
                    'valor': ev['valor'],
                    'juros': 0.0,
                    'dias_corridos': 0,
                    'taxa_efetiva': 0.0,
                    'incc': 0.0,
                    'ipca': 0.0,
                    'taxas_extra': [0.0] * len(taxas_extras),
                    'Total de mudança (R$)': 0.0,  # sem abatimento extra
                    'saldo': saldo                # saldo permanece inalterado
                })

            # parcela mensal
            mc_pre+=1; label=f"{mc_pre}ª Parcela Mensal Pré-Entrega"
            juros,dias,taxa_eff = tp.calculate(d_evt, saldo)
            incc = saldo * TAXA_INCC
            extras = sum(saldo*t['pct'] for t in taxas_extras if t['periodo']=='pré-entrega')
            abat = capacidade_pre - juros - incc - extras; saldo-=abat
            eventos.append({'data':d_evt,'parcela':'','tipo':label,'valor':capacidade_pre,
                            'juros':juros,'dias_corridos':dias,'taxa_efetiva':taxa_eff,
                            'incc':incc,'ipca':0.0,'taxas_extra':[0.0]*len(taxas_extras),
                            'Total de mudança (R$)':abat,'saldo':saldo})

            prev = cursor = d_evt + relativedelta(months=1)

        # ENTREGA
        for desc,v in [('Abatimento FGTS',fgts),('Abat. Banco',fin_banco)]:
            saldo-=v
            eventos.append({'data':data_entrega,'parcela':'','tipo':desc,'valor':v,
                            'juros':0.0,'dias_corridos':0,'taxa_efetiva':0.0,
                            'incc':0.0,'ipca':0.0,'taxas_extra':[0.0]*len(taxas_extras),
                            'Total de mudança (R$)':v,'saldo':saldo})
        eventos.append({'data':data_entrega,'parcela':'','tipo':'Chaves entregues','valor':0.0,
                        'juros':0.0,'dias_corridos':0,'taxa_efetiva':0.0,
                        'incc':0.0,'ipca':0.0,'taxas_extra':[0.0]*len(taxas_extras),
                        'Total de mudança (R$)':0.0,'saldo':saldo})

        # PÓS-ENTREGA
        tp2 = PaymentTracker(dia_pag, taxa_pos); tp2.last_date = data_entrega
        mc_pos = sc_pos = ac_pos = 0
        prev = cursor = data_entrega; parcelas = 1

        while saldo>0 and parcelas<=420:
            d_evt = adjust_day(cursor + relativedelta(months=1), dia_pag)

            # não-assoc
            for ev in [e for e in post_nr if not e['assoc'] and prev<e['data']<=d_evt]:
                if ev['tipo']=='Pagamento Semestral':
                    sc_pos+=1; label=f"{sc_pos}ª Parcela Semestral Pós-Entrega"; saldo-=ev['valor']
                elif ev['tipo']=='Pagamento Anual':
                    ac_pos+=1; label=f"{ac_pos}ª Parcela Anual Pós-Entrega"; saldo-=ev['valor']
                else:
                    mc_pos+=1; label=f"{mc_pos}ª Parcela Mensal Pós-Entrega"
                    juros,dias,taxa_eff = tp2.calculate(ev['data'], saldo)
                    ipca = saldo * TAXA_IPCA
                    extras = sum(saldo*t['pct'] for t in taxas_extras if t['periodo']=='pós-entrega')
                    abat = ev['valor'] - juros - ipca - extras; saldo-=abat
                eventos.append({**ev,'parcela':parcelas,'tipo':label,'juros':juros if 'juros' in locals() else 0.0,
                                'dias_corridos':dias if 'dias' in locals() else 0,'taxa_efetiva':taxa_eff if 'taxa_eff' in locals() else 0.0,
                                'incc':0.0,'ipca':ipca if 'ipca' in locals() else 0.0,
                                'taxas_extra':[0.0]*len(taxas_extras),'Total de mudança (R$)':ev['valor'],'saldo':saldo})

            # associados semestrais
            for ev in [e for e in sem_rec if e['assoc'] and prev<e['data']<=d_evt]:
                sc_pos+=1; label=f"{sc_pos}ª Parcela Semestral Pós-Entrega"; saldo-=ev['valor']
                eventos.append({**ev,'parcela':parcelas,'tipo':label,'juros':0.0,'dias_corridos':0,'taxa_efetiva':0.0,
                                'incc':0.0,'ipca':0.0,'taxas_extra':[0.0]*len(taxas_extras),'Total de mudança (R$)':ev['valor'],'saldo':saldo})

            # associados anuais
            for ev in [e for e in ann_rec if e['assoc'] and prev<e['data']<=d_evt]:
                ac_pos+=1; label=f"{ac_pos}ª Parcela Anual Pós-Entrega"; saldo-=ev['valor']
                eventos.append({**ev,'parcela':parcelas,'tipo':label,'juros':0.0,'dias_corridos':0,'taxa_efetiva':0.0,
                                'incc':0.0,'ipca':0.0,'taxas_extra':[0.0]*len(taxas_extras),'Total de mudança (R$)':ev['valor'],'saldo':saldo})
            # — Eventos não-recorrentes ASSOCIADOS pós‑entrega —
            for ev in [e for e in post_nr if e['assoc'] and prev < e['data'] <= d_evt]:
                eventos.append({
                    'data': ev['data'],
                    'parcela': parcelas,
                    'tipo': ev['tipo'],
                    'valor': ev['valor'],
                    'juros': 0.0,
                    'dias_corridos': 0,
                    'taxa_efetiva': 0.0,
                    'incc': 0.0,
                    'ipca': 0.0,
                    'taxas_extra': [0.0] * len(taxas_extras),
                    'Total de mudança (R$)': 0.0,
                    'saldo': saldo
                })


            # parcela mensal
            mc_pos+=1; label=f"{mc_pos}ª Parcela Mensal Pós-Entrega"
            juros,dias,taxa_eff = tp2.calculate(d_evt, saldo)
            ipca = saldo * TAXA_IPCA
            extras = sum(saldo*t['pct'] for t in taxas_extras if t['periodo']=='pós-entrega')
            abat = capacidade_pos - juros - ipca - extras; saldo-=abat
            eventos.append({'data':d_evt,'parcela':parcelas,'tipo':label,'valor':capacidade_pos,
                            'juros':juros,'dias_corridos':dias,'taxa_efetiva':taxa_eff,
                            'incc':0.0,'ipca':ipca,'taxas_extra':[0.0]*len(taxas_extras),
                            'Total de mudança (R$)':abat,'saldo':saldo})

            parcelas+=1; prev=cursor=d_evt

        # Monta Excel
# --- Monta e formata a planilha Excel ---
        wb = Workbook()
        ws = wb.active
        ws.title = f"Financ-{cliente}"[:31]

        # Cabeçalhos
        headers = ["Data", "Tipo", "Valor Pago (R$)"]
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = HEADER_FILL
            cell.font = Font(bold=True)

        # Linha inicial: saldo devedor inicial
        ws.append(["-", "-", valor_imovel])

        # Eventos (ordenados por data)
        for ev in sorted(eventos, key=lambda x: x['data']):
            ws.append([
                ev['data'],
                ev['tipo'],
                ev.get('valor', 0.0)
            ])

        # Linha em branco
        ws.append([''] * len(headers))

        # Totais
        soma_total = sum(ev['valor'] for ev in eventos if isinstance(ev['valor'], (int, float)))
        ws.append(['TOTAIS', '', soma_total])
        totals_row = ws.max_row
        ws.cell(row=totals_row, column=1).fill = HEADER_FILL
        ws.cell(row=totals_row, column=1).font = Font(bold=True)

        # Formatação de colunas
        for col_idx, h in enumerate(headers, start=1):
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if col_idx == 1:
                    cell.number_format = DATE_FORMAT
                else:
                    cell.number_format = CURRENCY_FORMAT

        # Ajuste automático de largura
        for col_cells in ws.columns:
            max_length = max(len(str(c.value)) for c in col_cells if c.value is not None)
            ws.column_dimensions[get_column_letter(col_cells[0].column)].width = max_length + 2

        # Download
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        st.download_button(
            "Download Excel",
            data=buf,
            file_name=f"Financiamento {cliente}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
if __name__ == "__main__":
    main()
