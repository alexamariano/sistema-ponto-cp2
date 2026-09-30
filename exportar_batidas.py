import sqlite3
import pandas as pd
import openpyxl
from datetime import datetime

# APONTA DIRETAMENTE PARA A PLANILHA ONDE OS RESULTADOS SÃO COLETADOS
MODELO_EXCEL = 'Sistema_de_Horas_Trabalhadas_DATATEMPO_CP2_conciliacao_automatica (1).xlsx'
NOME_BANCO = 'ponto.db'

def exportar_para_excel():
    print("Iniciando processamento e conciliação de horas na planilha oficial...")

    try:
        # 1. CARREGAR CADASTRO DE PESSOAS DO EXCEL
        df_pessoas = pd.read_excel(MODELO_EXCEL, sheet_name='Cadastro Pessoas', header=3)
        pessoas_map = {}
        for _, row in df_pessoas.dropna(subset=['ID']).iterrows():
            status = str(row['Status']).strip().upper() if pd.notna(row['Status']) else ''
            if status == 'ATIVO':
                pessoas_map[str(int(row['ID']))] = {
                    'base': str(row['Base']).strip() if pd.notna(row['Base']) else 'Hora',
                    'taxa_util': float(row['R$/h útil']) if pd.notna(row['R$/h útil']) else 0.0,
                    'taxa_fds': float(row['R$/h FDS']) if pd.notna(row['R$/h FDS']) else 0.0,
                    'diaria': float(row['R$/diária']) if pd.notna(row['R$/diária']) else 0.0
                }

        # 2. LER BATIDAS DO BANCO DE DADOS LOCAL (HTML)
        conn = sqlite3.connect(NOME_BANCO)
        df_batidas = pd.read_sql_query("SELECT * FROM batidas ORDER BY data_hora ASC", conn)
        conn.close()

        if df_batidas.empty:
            print("⚠️ Nenhuma batida encontrada no banco de dados.")
            return

        # 3. CARREGAR O MODELO DE CONCILIAÇÃO VIA OPENPYXL
        wb = openpyxl.load_workbook(MODELO_EXCEL)
        sheet_batidas = wb['Batidas']
        sheet_registros = wb['Registros']

        # Limpar linhas de dados anteriores a partir da linha 5 (sem deletar a estrutura da tabela XML)
        for r in range(5, sheet_batidas.max_row + 1):
            for c in range(1, 14):
                sheet_batidas.cell(row=r, column=c).value = None

        for r in range(5, sheet_registros.max_row + 1):
            for c in range(1, 20):
                sheet_registros.cell(row=r, column=c).value = None

        linha_batida = 5
        linha_registro = 5
        inicios_abertos = {}

        # 4. PROCESSAR REGRAS DO OFFICE SCRIPT EM PYTHON
        for _, batida in df_batidas.iterrows():
            b_id = str(batida['batida_id'])
            dt_obj = pd.to_datetime(batida['data_hora'])
            p_id = str(batida['pessoa_id'])
            nome = str(batida['nome'])
            tipo = str(batida['tipo']).strip().upper()
            projeto = str(batida['projeto'])
            atividade = str(batida['atividade'])
            observacao = str(batida.get('observacao', '')) if pd.notna(batida.get('observacao')) else ''
            
            data_serial = (dt_obj - datetime(1899, 12, 30)).total_seconds() / (24 * 3600)
            data_inteira = int(data_serial)
            
            dia_semana_num = dt_obj.weekday()
            dias_str = ["SEGUNDA-FEIRA", "TERÇA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA", "SEXTA-FEIRA", "SÁBADO", "DOMINGO"]
            dia_semana = dias_str[dia_semana_num]
            tipo_dia = "FDS/FERIADO" if dia_semana_num >= 5 else "DIA ÚTIL"
            inicio_vinculado = str(batida.get('inicio_vinculado', '')) if pd.notna(batida.get('inicio_vinculado')) else ''

            # Preencher aba "Batidas"
            sheet_batidas.cell(row=linha_batida, column=1, value=b_id)
            sheet_batidas.cell(row=linha_batida, column=2, value=dt_obj)
            sheet_batidas.cell(row=linha_batida, column=3, value=f'=IF(B{linha_batida}="","",INT(B{linha_batida}))')
            sheet_batidas.cell(row=linha_batida, column=4, value=dia_semana)
            sheet_batidas.cell(row=linha_batida, column=5, value=int(p_id))
            sheet_batidas.cell(row=linha_batida, column=6, value=nome)
            sheet_batidas.cell(row=linha_batida, column=7, value=tipo)
            sheet_batidas.cell(row=linha_batida, column=8, value=projeto)
            sheet_batidas.cell(row=linha_batida, column=9, value=atividade)
            sheet_batidas.cell(row=linha_batida, column=10, value=observacao)
            sheet_batidas.cell(row=linha_batida, column=11, value="Web HTML")
            sheet_batidas.cell(row=linha_batida, column=12, value=inicio_vinculado)
            sheet_batidas.cell(row=linha_batida, column=13, value="INÍCIO" if tipo == "INÍCIO" else "VINCULADO")
            linha_batida += 1

            # Preencher aba "Registros" (onde os resumos buscam as somas)
            if tipo == "INÍCIO":
                reg_id = "R" + b_id[1:]
                info_pessoa = pessoas_map.get(p_id, {'base': 'Hora', 'taxa_util': 0.0, 'taxa_fds': 0.0, 'diaria': 0.0})
                taxa = info_pessoa['taxa_util'] if tipo_dia == "DIA ÚTIL" else info_pessoa['taxa_fds']

                sheet_registros.cell(row=linha_registro, column=1, value=reg_id)
                sheet_registros.cell(row=linha_registro, column=2, value=dt_obj.date())
                sheet_registros.cell(row=linha_registro, column=3, value=dia_semana)
                sheet_registros.cell(row=linha_registro, column=4, value=tipo_dia)
                sheet_registros.cell(row=linha_registro, column=5, value=int(p_id))
                sheet_registros.cell(row=linha_registro, column=6, value=nome)
                sheet_registros.cell(row=linha_registro, column=7, value=projeto)
                sheet_registros.cell(row=linha_registro, column=8, value=atividade)
                sheet_registros.cell(row=linha_registro, column=9, value=dt_obj.strftime("%H:%M:%S"))
                sheet_registros.cell(row=linha_registro, column=13, value=info_pessoa['base'])
                sheet_registros.cell(row=linha_registro, column=14, value=taxa)
                sheet_registros.cell(row=linha_registro, column=15, value=info_pessoa['diaria'])
                sheet_registros.cell(row=linha_registro, column=17, value="EM ABERTO")
                sheet_registros.cell(row=linha_registro, column=18, value=1)
                sheet_registros.cell(row=linha_registro, column=19, value=dt_obj.strftime("%Y-%m"))

                inicios_abertos[b_id] = {'linha': linha_registro, 'data_inicio': dt_obj, 'taxa': taxa, 'base': info_pessoa['base'], 'diaria': info_pessoa['diaria']}
                linha_registro += 1

            elif tipo == "FIM":
                chave_inicio = inicio_vinculado
                if not chave_inicio:
                    for k, v in reversed(list(inicios_abertos.items())):
                        if k.endswith(f"-{p_id}"):
                            chave_inicio = k
                            break

                if chave_inicio in inicios_abertos:
                    dados_inicio = inicios_abertos[chave_inicio]
                    linha_alvo = dados_inicio['linha']

                    sheet_registros.cell(row=linha_alvo, column=10, value=dt_obj.strftime("%H:%M:%S"))
                    horas = max(0.0, (dt_obj - dados_inicio['data_inicio']).total_seconds() / 3600.0)
                    sheet_registros.cell(row=linha_alvo, column=11, value=round(horas, 2))
                    sheet_registros.cell(row=linha_alvo, column=12, value=observacao)

                    valor = dados_inicio['diaria'] if dados_inicio['base'].upper() == "DIÁRIA" else (horas * dados_inicio['taxa'])
                    sheet_registros.cell(row=linha_alvo, column=16, value=round(valor, 2))
                    sheet_registros.cell(row=linha_alvo, column=17, value="CONCLUÍDO")

                    del inicios_abertos[chave_inicio]

        # Salvar na planilha oficial de conciliação
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nome_saida = f'Sistema_Conciliacao_Final_{timestamp}.xlsx'
        wb.save(nome_saida)

        print("=" * 65)
        print("✅ CONCILIAÇÃO AUTOMÁTICA GERADA COM SUCESSO!")
        print(f"📁 Planilha salva como: {nome_saida}")
        print("=" * 65)

    except Exception as e:
        print(f"❌ Ocorreu um erro: {e}")

if __name__ == '__main__':
    exportar_para_excel()