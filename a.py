from flask import Flask, render_template, request, jsonify
import pandas as pd
import os

app = Flask(__name__)

# Configuração para carregar a planilha
PLANILHA_PATH = 'clientes.xlsx'  # Coloque o caminho da sua planilha aqui
df = None

def carregar_planilha():
    """Carrega a planilha Excel"""
    global df
    try:
        if os.path.exists(PLANILHA_PATH):
            df = pd.read_excel(PLANILHA_PATH)
            # Limpar espaços em branco nas colunas de texto
            for col in ['APELIDO', 'REGIAO', 'NOME']:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()
            print(f"Planilha carregada com sucesso! {len(df)} registros encontrados.")
            print("Colunas disponíveis:", df.columns.tolist())
        else:
            print(f"Arquivo {PLANILHA_PATH} não encontrado!")
            df = pd.DataFrame()  # DataFrame vazio
    except Exception as e:
        print(f"Erro ao carregar planilha: {e}")
        df = pd.DataFrame()

@app.route('/')
def index():
    """Página principal"""
    return render_template('index.html')

@app.route('/buscar', methods=['POST'])
def buscar_cliente():
    """Busca cliente por nome ou ID"""
    global df
    
    if df is None or df.empty:
        return jsonify({
            'sucesso': False, 
            'erro': 'Planilha não carregada ou vazia'
        })
    
    termo_busca = request.json.get('termo', '').strip()
    
    if not termo_busca:
        return jsonify({
            'sucesso': False, 
            'erro': 'Termo de busca não pode estar vazio'
        })
    
    try:
        # Buscar por NOME (busca parcial, case insensitive)
        resultado_nome = df[df['NOME'].str.contains(termo_busca, case=False, na=False)]
        
        # Buscar por ID_CLIENTE (busca exata)
        try:
            # Tentar converter o termo para número para busca por ID
            id_busca = int(termo_busca)
            resultado_id = df[df['ID_CLIENTE'] == id_busca]
        except ValueError:
            # Se não conseguir converter para número, não busca por ID
            resultado_id = pd.DataFrame()
        
        # Combinar resultados (remove duplicatas)
        resultado_final = pd.concat([resultado_nome, resultado_id]).drop_duplicates()
        
        if resultado_final.empty:
            return jsonify({
                'sucesso': False, 
                'erro': f'Nenhum cliente encontrado para "{termo_busca}"'
            })
        
        # Converter resultado para lista de dicionários
        clientes = []
        for _, row in resultado_final.iterrows():
            cliente = {
                'apelido': row.get('APELIDO', ''),
                'regiao': row.get('REGIAO', ''),
                'nome': row.get('NOME', ''),
                'id_cliente': row.get('ID_CLIENTE', ''),
                'pontos': row.get('SALDO', 0),
                'pontos_resgatados': row.get('SALDO RESGATADO EM TODO PERIODO', 0),
                'cashback': row.get('CASHBACK', 0),
                'cashback_resgatado': row.get('CASHBACK RESGATADO', 0)
            }
            clientes.append(cliente)
        
        return jsonify({
            'sucesso': True, 
            'clientes': clientes,
            'total': len(clientes)
        })
        
    except Exception as e:
        return jsonify({
            'sucesso': False, 
            'erro': f'Erro na busca: {str(e)}'
        })

@app.route('/recarregar')
def recarregar_planilha():
    """Recarrega a planilha"""
    carregar_planilha()
    total_registros = len(df) if df is not None else 0
    return jsonify({
        'sucesso': True, 
        'mensagem': f'Planilha recarregada com {total_registros} registros'
    })

if __name__ == '__main__':
    # Carregar planilha na inicialização
    carregar_planilha()
    
    # Executar aplicação
    print("Iniciando aplicação...")
    print("Acesse: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)