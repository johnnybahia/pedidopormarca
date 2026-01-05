# Exemplo de script para enviar dados para a planilha

import requests
import json

# Substitua pela URL do seu Web App
URL_WEBAPP = "https://script.google.com/macros/s/SEU_ID_AQUI/exec"

def enviar_pedido():
    # Dados do pedido
    pedido = {
        "dataPedido": "05/01/2026",        # Data do pedido (DD/MM/AAAA)
        "dataRecebimento": "10/01/2026",   # Data de recebimento (DD/MM/AAAA)
        "arquivo": "pedido_001.pdf",
        "cliente": "Loja ABC",
        "marca": "Nike",
        "local": "Salvador",
        "qtd": 100,
        "unidade": "PAR",
        "valor": "R$ 15.000,00"
    }
    
    # Wrapper obrigatório
    dados = {"pedidos": [pedido]}
    
    print("Enviando dados:")
    print(json.dumps(dados, indent=2, ensure_ascii=False))
    
    try:
        response = requests.post(
            URL_WEBAPP,
            json=dados,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        print(f"\nStatus Code: {response.status_code}")
        print(f"Resposta: {response.text}")
        
        if response.status_code == 200:
            resultado = response.json()
            print(f"\n✅ Sucesso: {resultado}")
        else:
            print(f"\n❌ Erro HTTP: {response.status_code}")
            
    except Exception as e:
        print(f"\n❌ Erro: {e}")

if __name__ == "__main__":
    enviar_pedido()
