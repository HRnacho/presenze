from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from datetime import datetime, timedelta
import json
import os
import calendar
from functools import wraps

app = Flask(__name__)
app.secret_key = 'direzionelavoro-presenze-2025-secret-key'

DATA_FILE = os.path.join('data', 'presenze.json')

USERS = {
    'gianluca': {'password': 'direzione2025', 'nome': 'Gianluca Bittoni'},
    'ignacio': {'password': 'direzione2025', 'nome': 'Ignacio Sorcaburu Ciglieri'},
    'simone': {'password': 'direzione2025', 'nome': 'Simone Mascellari'}
}

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_data(data):
    os.makedirs('data', exist_ok=True)
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

@app.route('/')
def index():
    if 'username' not in session:
        return redirect(url_for('login'))
    return redirect(url_for('calendario'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username in USERS and USERS[username]['password'] == password:
            session['username'] = username
            session['nome_completo'] = USERS[username]['nome']
            return redirect(url_for('calendario'))
        else:
            return render_template('login.html', error='Credenziali non valide')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/calendario')
@login_required
def calendario():
    now = datetime.now()
    year = int(request.args.get('year', now.year))
    month = int(request.args.get('month', now.month))
    data = load_data()
    utenti = [{'username': u, 'nome': i['nome']} for u, i in USERS.items()]
    return render_template('calendario.html', year=year, month=month, utenti=utenti,
                         current_user=session['username'], nome_completo=session['nome_completo'])

@app.route('/api/presenze/<year>/<month>')
@login_required
def get_presenze(year, month):
    data = load_data()
    key = f"{year}-{month.zfill(2)}"
    return jsonify(data.get(key, {}))

@app.route('/api/presenza', methods=['POST'])
@login_required
def salva_presenza():
    try:
        info = request.json
        data = load_data()
        date_obj = datetime.strptime(info['date'], '%Y-%m-%d')
        key = date_obj.strftime('%Y-%m')
        
        if key not in data:
            data[key] = {}
        if info['username'] not in data[key]:
            data[key][info['username']] = {}
        
        data[key][info['username']][info['date']] = {
            'tipo': info['tipo'],
            'ore_lavorate': float(info.get('ore_lavorate', 0)),
            'ore_assenza': float(info.get('ore_assenza', 0)),
            'note': info.get('note', ''),
            'updated_at': datetime.now().isoformat()
        }
        
        save_data(data)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 400

@app.route('/api/elimina_presenza', methods=['POST'])
@login_required
def elimina_presenza():
    try:
        info = request.json
        data = load_data()
        date_obj = datetime.strptime(info['date'], '%Y-%m-%d')
        key = date_obj.strftime('%Y-%m')
        if key in data and info['username'] in data[key] and info['date'] in data[key][info['username']]:
            del data[key][info['username']][info['date']]
            save_data(data)
            return jsonify({'success': True})
        return jsonify({'success': False, 'error': 'Non trovata'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 400

@app.route('/esporta/<year>/<month>')
@login_required
def esporta_excel(year, month):
    from io import BytesIO
    from flask import send_file
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    data = load_data()
    key = f"{year}-{month.zfill(2)}"
    presenze_mese = data.get(key, {})
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Presenze"
    
    month_name = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
                  'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'][int(month)-1]
    
    ws['A1'] = f"FOGLIO PRESENZE - {month_name.upper()} {year}"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:AG1')
    
    ws['A3'] = "COGNOME E NOME"
    ws['A3'].font = Font(bold=True, size=11)
    ws['A3'].fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    ws['B3'] = ""
    ws['B3'].fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    
    num_days = calendar.monthrange(int(year), int(month))[1]
    for day in range(1, num_days + 1):
        col = day + 2
        ws.cell(row=3, column=col).value = day
        ws.cell(row=3, column=col).font = Font(bold=True, size=10)
        ws.cell(row=3, column=col).alignment = Alignment(horizontal='center')
        ws.cell(row=3, column=col).fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 4
    
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 6
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    start_row = 4
    
    for idx, (username, user_info) in enumerate(USERS.items()):
        nome_completo = user_info['nome']
        presenze_user = presenze_mese.get(username, {})
        
        base_row = start_row + (idx * 4)
        
        ws.merge_cells(f'A{base_row}:A{base_row+3}')
        ws.cell(row=base_row, column=1).value = nome_completo
        ws.cell(row=base_row, column=1).font = Font(bold=True, size=10)
        ws.cell(row=base_row, column=1).alignment = Alignment(vertical='center', horizontal='left')
        ws.cell(row=base_row, column=1).border = thin_border
        
        labels = ['ORD', 'STR', 'ASS', 'GIUST']
        for i, label in enumerate(labels):
            ws.cell(row=base_row + i, column=2).value = label
            ws.cell(row=base_row + i, column=2).font = Font(bold=True, size=9)
            ws.cell(row=base_row + i, column=2).alignment = Alignment(horizontal='center')
            ws.cell(row=base_row + i, column=2).border = thin_border
        
        for day in range(1, num_days + 1):
            date_str = f"{year}-{month.zfill(2)}-{str(day).zfill(2)}"
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            col = day + 2
            
            is_weekend = date_obj.weekday() in [5, 6]
            
            if is_weekend:
                gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                for i in range(4):
                    cell = ws.cell(row=base_row + i, column=col)
                    cell.fill = gray_fill
                    cell.border = thin_border
            else:
                if date_str in presenze_user:
                    presenza = presenze_user[date_str]
                    tipo = presenza.get('tipo')
                    ore_lavorate = presenza.get('ore_lavorate', 0)
                    ore_assenza = presenza.get('ore_assenza', 0)
                    
                    cell_ord = ws.cell(row=base_row, column=col)
                    if ore_lavorate > 0:
                        cell_ord.value = ore_lavorate
                        cell_ord.alignment = Alignment(horizontal='center')
                    cell_ord.border = thin_border
                    
                    cell_str = ws.cell(row=base_row + 1, column=col)
                    cell_str.border = thin_border
                    
                    cell_ass = ws.cell(row=base_row + 2, column=col)
                    if ore_assenza > 0:
                        cell_ass.value = ore_assenza
                        cell_ass.alignment = Alignment(horizontal='center')
                    cell_ass.border = thin_border
                    
                    cell_giust = ws.cell(row=base_row + 3, column=col)
                    if tipo != 'presenza':
                        codice = ''
                        if tipo == 'ferie':
                            codice = 'F'
                        elif tipo == 'rol':
                            codice = 'R'
                        elif tipo == 'malattia':
                            codice = 'M'
                        elif tipo == 'permesso':
                            codice = 'P'
                        
                        if codice:
                            cell_giust.value = codice
                            cell_giust.alignment = Alignment(horizontal='center')
                            cell_giust.font = Font(bold=True)
                    cell_giust.border = thin_border
                else:
                    for i in range(4):
                        cell = ws.cell(row=base_row + i, column=col)
                        cell.border = thin_border
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"Presenze_{month_name}_{year}.xlsx"
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=filename)

if __name__ == '__main__':
    os.makedirs('data', exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000)
