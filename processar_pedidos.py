import pdfplumber
import re
import os
import shutil
import requests
import json
from datetime import datetime

# ================= CONFIGURAÇÃO =================
URL_WEBAPP = "https://script.google.com/macros/s/AKfycbwYOBCSEwak53AA_eXIzubyi5dHSrKL2wpAeKvogzR3MHaxIsPDOuFuRaxTL3WwLHNX/exec"

PASTA_ENTRADA = './pedidos'
PASTA_LIDOS = './pedidos/lidos'
# =================================================

def limpar_valor_monetario(texto):
    if not texto: return 0.0
    texto = texto.lower().replace('r$', '').strip().replace('.', '').replace(',', '.')
    try: return float(texto)
    except: return 0.0

def identificar_unidade(texto):
    if re.search(r'\bPAR\b', texto, re.IGNORECASE): return "PAR"
    if re.search(r'\bM\b|\bMTS\b|\bMETRO\b', texto, re.IGNORECASE): return "METRO"
    return "UNID"

def extrair_local_entrega(texto):
    texto_upper = texto.upper()
    if "NE-03" in texto_upper or "SEST" in texto_upper: return "Santo Estêvão (NE-03)"
    if "NE-08" in texto_upper or "ITABERABA" in texto_upper: return "Itaberaba (NE-08)"
    if "NE-09" in texto_upper or "VDC" in texto_upper: return "Vitória da Conquista (NE-09)"

    matches = re.findall(r'Cidade:\s*([A-Z\s]+)', texto)
    for c in matches:
        if "CRUZ DAS ALMAS" not in c.upper(): return c.strip().upper()
    return "Local Não Identificado"

def processar_pdf_dass(caminho_arquivo, nome_arquivo):
    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += page.extract_text() or ""

            if "DASS" not in texto_completo and "01287588" not in texto_completo:
                return None

            # --- 1. DATA DE RECEBIMENTO (Baseado na Data da Emissão) ---
            match_emissao = re.search(r'Data da emissão:\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.IGNORECASE)
            if match_emissao:
                data_recebimento = match_emissao.group(1)
            else:
                match_header = re.search(r'Hora.*?Data\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)
                data_recebimento = match_header.group(1) if match_header else datetime.now().strftime("%d/%m/%Y")

            # --- DADOS GERAIS DO PEDIDO ---
            match_marca = re.search(r'Marca:\s*([^\n]+)', texto_completo)
            marca_geral = match_marca.group(1).strip() if match_marca else "N/D"

            local_geral = extrair_local_entrega(texto_completo)

            # Captura Totais Gerais
            valor_total_doc = 0.0
            qtd_total_doc = 0

            match_valor_tot = re.search(r'Total valor:\s*([\d\.,]+)', texto_completo)
            if match_valor_tot: valor_total_doc = limpar_valor_monetario(match_valor_tot.group(1))

            match_qtd_tot = re.search(r'Total peças:\s*([\d\.,]+)', texto_completo)
            if match_qtd_tot: qtd_total_doc = int(limpar_valor_monetario(match_qtd_tot.group(1)))

            # --- 2. SEPARAÇÃO EM PARTES (SPLIT) ---
            itens_encontrados = re.findall(r'(\d{8}).*?(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)

            lista_pedidos_extraidos = []

            if not itens_encontrados:
                # Se não achou itens individuais pelo NCM, cria UM pedido genérico com o total
                lista_pedidos_extraidos.append({
                    "dataEntrega": data_recebimento,     # ✅ camelCase - Fallback
                    "dataRecebimento": data_recebimento, # ✅ camelCase
                    "arquivo": nome_arquivo,
                    "cliente": "Grupo DASS",
                    "marca": marca_geral,
                    "local": local_geral,
                    "qtd": qtd_total_doc,
                    "unidade": identificar_unidade(texto_completo),
                    "valor_raw": valor_total_doc  # Valor numérico cru
                })
            else:
                # Se achou itens, cria uma linha para cada um
                for i, item in enumerate(itens_encontrados):
                    ncm_code = item[0]
                    data_entrega = item[1]

                    # LÓGICA DE VALOR:
                    # atribuímos o Valor Total apenas à PRIMEIRA linha para não duplicar o faturamento.
                    # Nas outras linhas colocamos 0, mas mantemos a Data de Entrega para controle logístico.
                    val_linha = valor_total_doc if i == 0 else 0.0
                    qtd_linha = qtd_total_doc if i == 0 else 0

                    lista_pedidos_extraidos.append({
                        "dataEntrega": data_entrega,         # ✅ camelCase - A data específica deste item
                        "dataRecebimento": data_recebimento, # ✅ camelCase - A data de emissão do documento
                        "arquivo": f"{nome_arquivo} ({i+1})",  # Nome do arquivo com sufixo (1), (2)...
                        "cliente": "Grupo DASS",
                        "marca": marca_geral,
                        "local": local_geral,
                        "qtd": qtd_linha,
                        "unidade": identificar_unidade(texto_completo),
                        "valor_raw": val_linha
                    })

            # Formata os valores finais para String (R$)
            resultados_finais = []
            for p in lista_pedidos_extraidos:
                p["valor"] = f"R$ {p['valor_raw']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                del p["valor_raw"]
                resultados_finais.append(p)

            return resultados_finais

    except Exception as e:
        print(f"Erro ao abrir {nome_arquivo}: {e}")
        return []

def mover_arquivos_processados(lista_arquivos):
    if not os.path.exists(PASTA_LIDOS): os.makedirs(PASTA_LIDOS)
    print(f"\n📦 Movendo arquivos processados para: {PASTA_LIDOS}")
    for arquivo in set(lista_arquivos):
        try:
            caminho_origem = os.path.join(PASTA_ENTRADA, arquivo)
            caminho_destino = os.path.join(PASTA_LIDOS, arquivo)
            if os.path.exists(caminho_destino): os.remove(caminho_destino)
            shutil.move(caminho_origem, caminho_destino)
        except: pass

def main():
    if not os.path.exists(PASTA_ENTRADA):
        os.makedirs(PASTA_ENTRADA)
        print(f"Pasta criada. Coloque PDFs em '{PASTA_ENTRADA}'.")
        return

    todos_pedidos_para_envio = []
    arquivos_para_mover = []

    arquivos = [f for f in os.listdir(PASTA_ENTRADA) if f.lower().endswith('.pdf')]

    print(f"📂 Processando {len(arquivos)} arquivos...")
    print("-" * 75)
    print(f"{'RECEBIMENTO':<12} | {'ENTREGA':<12} | {'MARCA':<15} | {'VALOR'}")
    print("-" * 75)

    for arq in arquivos:
        lista_pedidos = processar_pdf_dass(os.path.join(PASTA_ENTRADA, arq), arq)

        if lista_pedidos:
            for p in lista_pedidos:
                todos_pedidos_para_envio.append(p)
                print(f"✅ {p['dataRecebimento']:<12} | {p['dataEntrega']:<12} | {p['marca']:<15} | {p['valor']}")
            arquivos_para_mover.append(arq)
        else:
            print(f"⚠️ Ignorado: {arq}")

    if todos_pedidos_para_envio:
        print("-" * 75)
        print(f"📤 Enviando {len(todos_pedidos_para_envio)} linhas para Google Sheets...")
        print(f"\n📋 JSON enviado:")
        print(json.dumps({"pedidos": todos_pedidos_para_envio}, indent=2, ensure_ascii=False))

        try:
            response = requests.post(URL_WEBAPP, json={"pedidos": todos_pedidos_para_envio}, timeout=30)

            print(f"\n📡 Status: {response.status_code}")
            print(f"📡 Resposta: {response.text}")

            if response.status_code == 200:
                print(f"\n☁️ SUCESSO! Google recebeu os dados.")
                mover_arquivos_processados(arquivos_para_mover)
            else:
                print(f"\n❌ Erro HTTP {response.status_code}")

        except Exception as e:
            print(f"\n❌ Erro de conexão: {e}")
    else:
        print("\n⚠️ Nenhum pedido válido encontrado.")

    input("\nPressione ENTER para fechar...")

if __name__ == "__main__":
    main()
