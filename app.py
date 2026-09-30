import sqlite3
from flask import Flask, request, jsonify, render_template_string
from datetime import datetime

app = Flask(__name__)

# --- 1. BASE DE DADOS CADASTRAIS (Extraídos do Excel) ---
PESSOAS = {
    "1001": "Adriana Pereira dos Santos",
    "1002": "Alex de Almeida Mariano",
    "1003": "Amanda do Carmo Ribeiro",
    "1004": "Angela Bertoli",
    "1005": "Carolina Fantini Vidigal Diniz Dias",
    "1006": "Eliane dos Santos Oliveira",
    "1007": "Fernanda Lopes",
    "1008": "Flávia Reis",
    "1009": "Junia Maria Santos",
    "1010": "Nataly Tayê",
    "1011": "Nathalia Arruda",
    "1012": "Nelma Lucia dos S. B. Brandão",
    "1013": "Rafaela Maria Resende Lara",
    "1014": "Robert Filipe Orlando de Souza",
    "1015": "Vitor Guilherme Miguel Rocha",
    "1016": "Ana Luisa Nardin",
    "1017": "Renato Inácio da Silva"
}

PROJETOS = [
    {"codigo": "PRJ-001", "nome": "BC - CONFIANÇA PIX"},
    {"codigo": "PRJ-002", "nome": "CFI / ACCION"},
    {"codigo": "PRJ-003", "nome": "SICOOB CREDICOM"},
    {"codigo": "PRJ-004", "nome": "OTEMPO"},
    {"codigo": "PRJ-005", "nome": "ESTADUAL - OTEMPO"},
    {"codigo": "PRJ-006", "nome": "CDL/BH"},
    {"codigo": "PRJ-007", "nome": "TOOLKIT"},
    {"codigo": "PRJ-008", "nome": "CAMPANHA"},
    {"codigo": "PRJ-009", "nome": "ADMINISTRATIVO"},
    {"codigo": "PRJ-010", "nome": "OUTROS"},
    {"codigo": "PRJ-011", "nome": "BETIM - TRACKING"}
]

ATIVIDADES = [
    {"codigo": "ATV-001", "nome": "SUPERVISÃO"},
    {"codigo": "ATV-002", "nome": "CHECAGEM"},
    {"codigo": "ATV-003", "nome": "CODIFICAÇÃO"},
    {"codigo": "ATV-004", "nome": "ANÁLISE"},
    {"codigo": "ATV-005", "nome": "RECRUTAMENTO"},
    {"codigo": "ATV-006", "nome": "APLICAÇÃO / ENTREVISTA"},
    {"codigo": "ATV-007", "nome": "TRATAMENTO DE DADOS"},
    {"codigo": "ATV-008", "nome": "RELATÓRIO"},
    {"codigo": "ATV-009", "nome": "ADMINISTRATIVO"},
    {"codigo": "ATV-010", "nome": "APOIO OPERACIONAL"},
    {"codigo": "ATV-011", "nome": "TRANSCRIÇÃO"},
    {"codigo": "ATV-012", "nome": "MODERAÇÃO"}
]

# --- 2. INICIALIZAÇÃO DO BANCO DE DADOS DE BATIDAS ---
def init_db():
    conn = sqlite3.connect('ponto.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS batidas (
            batida_id TEXT PRIMARY KEY,
            data_hora TEXT,
            data TEXT,
            dia_semana TEXT,
            pessoa_id TEXT,
            nome TEXT,
            tipo TEXT,
            projeto TEXT,
            atividade TEXT,
            observacao TEXT,
            origem TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# --- 3. INTERFACE HTML ---
HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Batida de Ponto — DATATEMPO/CP2</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f0f2f5; color: #333; }
        .container { max-width: 480px; margin: 20px auto; background: #ffffff; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
        h2 { text-align: center; color: #1a365d; margin-top: 0; font-size: 1.4em; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; }
        .form-group { margin-bottom: 18px; }
        label { display: block; margin-bottom: 6px; font-weight: 600; font-size: 0.9em; color: #4a5568; }
        input, select { width: 100%; padding: 11px; box-sizing: border-box; border: 1px solid #cbd5e0; border-radius: 6px; font-size: 0.95em; }
        input:focus, select:focus { outline: none; border-color: #3182ce; box-shadow: 0 0 0 3px rgba(49, 130, 206, 0.15); }
        .btn-group { display: flex; gap: 10px; margin-top: 10px; }
        .btn { flex: 1; padding: 12px; border: none; border-radius: 6px; color: white; cursor: pointer; font-weight: bold; font-size: 1em; transition: background 0.2s; }
        .btn-inicio { background-color: #28a745; }
        .btn-inicio:hover { background-color: #218838; }
        .btn-fim { background-color: #dc3545; }
        .btn-fim:hover { background-color: #c82333; }
        .btn-espelho { background-color: #3182ce; width: 100%; margin-top: 20px; }
        .btn-espelho:hover { background-color: #2b6cb0; }
        .extrato { margin-top: 25px; display: none; background: #faf5ff; padding: 15px; border-radius: 8px; border: 1px solid #e9d8fd; }
        .extrato h3 { margin-top: 0; font-size: 1.05em; color: #553c9a; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 0.85em; }
        th, td { border: 1px solid #e2e8f0; padding: 8px; text-align: left; }
        th { background-color: #edf2f7; color: #2d3748; }
        .badge-inicio { color: #276749; font-weight: bold; }
        .badge-fim { color: #9b2c2c; font-weight: bold; }
    </style>
</head>
<body>
    <div class="container">
        <h2>DATATEMPO / CP2<br><span style="font-size: 0.7em; font-weight: normal; color: #718096;">Registro Individual de Ponto</span></h2>
        
        <div class="form-group">
            <label for="pessoa_select">Selecione seu Nome/ID:</label>
            <select id="pessoa_select" onchange="atualizarId()">
                <option value="">-- Selecione na lista --</option>
                {% for id, nome in pessoas.items() %}
                    <option value="{{ id }}">{{ id }} - {{ nome }}</option>
                {% endfor %}
            </select>
        </div>

        <div class="form-group">
            <label for="pessoa_id">Ou digite diretamente seu ID:</label>
            <input type="text" id="pessoa_id" placeholder="Ex: 1002" oninput="atualizarSelect()">
        </div>

        <div class="form-group">
            <label for="projeto">Projeto:</label>
            <select id="projeto">
                {% for p in projetos %}
                    <option value="{{ p.nome }}">{{ p.codigo }} - {{ p.nome }}</option>
                {% endfor %}
            </select>
        </div>

        <div class="form-group">
            <label for="atividade">Atividade:</label>
            <select id="atividade">
                {% for a in atividades %}
                    <option value="{{ a.nome }}">{{ a.codigo }} - {{ a.nome }}</option>
                {% endfor %}
            </select>
        </div>

        <div class="form-group">
            <label for="observacao">Observação (Opcional):</label>
            <input type="text" id="observacao" placeholder="Ex: Checagem de amostragem no bairro X">
        </div>

        <div class="btn-group">
            <button class="btn btn-inicio" onclick="registrar('INÍCIO')">INÍCIO</button>
            <button class="btn btn-fim" onclick="registrar('FIM')">FIM</button>
        </div>

        <button class="btn btn-espelho" onclick="carregarEspelho()">Ver Meu Espelho de Hoje</button>

        <div class="extrato" id="extrato">
            <h3 id="titulo-espelho">Espelho do Ponto</h3>
            <table>
                <thead>
                    <tr>
                        <th>Hora</th>
                        <th>Tipo</th>
                        <th>Projeto</th>
                        <th>Atividade</th>
                    </tr>
                </thead>
                <tbody id="tabela-espelho"></tbody>
            </table>
        </div>
    </div>

    <script>
        function atualizarId() {
            const select = document.getElementById('pessoa_select');
            document.getElementById('pessoa_id').value = select.value;
        }

        function atualizarSelect() {
            const idInput = document.getElementById('pessoa_id').value;
            const select = document.getElementById('pessoa_select');
            select.value = idInput;
        }

        async function registrar(tipo) {
            const pessoa_id = document.getElementById('pessoa_id').value;
            if (!pessoa_id) {
                alert("Por favor, selecione ou digite o seu ID de funcionário.");
                return;
            }

            const payload = {
                pessoa_id: pessoa_id,
                projeto: document.getElementById('projeto').value,
                atividade: document.getElementById('atividade').value,
                observacao: document.getElementById('observacao').value,
                tipo: tipo
            };

            const response = await fetch('/api/bater_ponto', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const result = await response.json();
            if (response.ok) {
                alert(result.mensagem);
                document.getElementById('observacao').value = '';
                carregarEspelho();
            } else {
                alert("Erro: " + result.erro);
            }
        }

        async function carregarEspelho() {
            const pessoa_id = document.getElementById('pessoa_id').value;
            if (!pessoa_id) {
                alert("Selecione seu ID para carregar seu espelho de ponto.");
                return;
            }

            const response = await fetch('/api/espelho?id=' + pessoa_id);
            const dados = await response.json();

            if (!response.ok) {
                alert("Erro ao buscar espelho: " + dados.erro);
                return;
            }

            const tbody = document.getElementById('tabela-espelho');
            tbody.innerHTML = '';

            dados.registros.forEach(b => {
                const tr = document.createElement('tr');
                const hora = b.data_hora.split(' ')[1];
                const badgeClass = b.tipo === 'INÍCIO' ? 'badge-inicio' : 'badge-fim';
                
                tr.innerHTML = `
                    <td>${hora}</td>
                    <td class="${badgeClass}">${b.tipo}</td>
                    <td>${b.projeto}</td>
                    <td>${b.atividade}</td>
                `;
                tbody.appendChild(tr);
            });

            document.getElementById('titulo-espelho').innerText = `Espelho de Hoje - ${dados.nome}`;
            document.getElementById('extrato').style.display = 'block';
        }
    </script>
</body>
</html>
"""

# --- 4. ROTAS DO SERVIDOR ---
@app.route('/')
def index():
    return render_template_string(HTML, pessoas=PESSOAS, projetos=PROJETOS, atividades=ATIVIDADES)

@app.route('/api/bater_ponto', methods=['POST'])
def bater_ponto():
    dados = request.json
    pessoa_id = str(dados.get('pessoa_id', '')).strip()
    
    # Valida se a pessoa existe no cadastro
    if pessoa_id not in PESSOAS:
        return jsonify({"erro": f"ID {pessoa_id} não encontrado no cadastro de pessoas!"}), 400

    nome_pessoa = PESSOAS[pessoa_id]
    agora = datetime.now()
    
    data_hora = agora.strftime('%Y-%m-%d %H:%M:%S')
    data = agora.strftime('%Y-%m-%d')
    dias_pt = ['SEGUNDA-FEIRA', 'TERÇA-FEIRA', 'QUARTA-FEIRA', 'QUINTA-FEIRA', 'SEXTA-FEIRA', 'SÁBADO', 'DOMINGO']
    dia_semana = dias_pt[agora.weekday()]
    
    timestamp_id = agora.strftime('%Y%m%d%H%M%S')
    batida_id = f"B{timestamp_id}-{pessoa_id}"
    
    tipo = dados.get('tipo')
    projeto = dados.get('projeto')
    atividade = dados.get('atividade')
    observacao = dados.get('observacao', '')
    origem = "Web HTML"

    conn = sqlite3.connect('ponto.db')
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO batidas (batida_id, data_hora, data, dia_semana, pessoa_id, nome, tipo, projeto, atividade, observacao, origem)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (batida_id, data_hora, data, dia_semana, pessoa_id, nome_pessoa, tipo, projeto, atividade, observacao, origem))
    conn.commit()
    conn.close()

    return jsonify({"mensagem": f"Ponto de {tipo} registrado com sucesso para {nome_pessoa} às {agora.strftime('%H:%M:%S')}!"})

@app.route('/api/espelho', methods=['GET'])
def get_espelho():
    pessoa_id = str(request.args.get('id', '')).strip()
    
    if pessoa_id not in PESSOAS:
        return jsonify({"erro": "ID inválido."}), 400

    data_hoje = datetime.now().strftime('%Y-%m-%d')
    
    conn = sqlite3.connect('ponto.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute('''
        SELECT data_hora, tipo, projeto, atividade, observacao 
        FROM batidas 
        WHERE pessoa_id = ? AND data = ?
        ORDER BY data_hora ASC
    ''', (pessoa_id, data_hoje))
    registros = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return jsonify({
        "nome": PESSOAS[pessoa_id],
        "registros": registros
    })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)