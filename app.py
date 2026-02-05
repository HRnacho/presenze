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
            'ore_giustificativo': float(info.get('ore_giustificativo', 0)),
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
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    weekend_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    
    for idx, (username, user_info) in enumerate(USERS.items()):
        ws = wb.active if idx == 0 else wb.create_sheet(title=user_info['nome'].split()[0])
        if idx == 0:
            ws.title = user_info['nome'].split()[0]
        
        ws['A1'] = f"FOGLIO PRESENZE - {user_info['nome']}"
        ws['A1'].font = Font(bold=True, size=14)
        ws.merge_cells('A1:H1')
        ws['A2'] = f"Mese: {calendar.month_name[int(month)]} {year}"
        ws['A2'].font = Font(bold=True, size=12)
        ws.merge_cells('A2:H2')
        
        headers = ['Data', 'Giorno', 'Ore Ordinarie', 'Ferie', 'ROL', 'Permessi Retribuiti', 'Malattia', 'Note']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center')
        
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 18
        ws.column_dimensions['G'].width = 12
        ws.column_dimensions['H'].width = 30
        
        num_days = calendar.monthrange(int(year), int(month))[1]
        presenze_user = presenze_mese.get(username, {})
        totali = {'ore_ordinarie': 0, 'ferie': 0, 'rol': 0, 'permessi': 0, 'malattia': 0}
        
        row = 5
        for day in range(1, num_days + 1):
            date_str = f"{year}-{month.zfill(2)}-{str(day).zfill(2)}"
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            day_name = ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica'][date_obj.weekday()]
            
            ore_ordinarie = ferie = rol = permessi = malattia = 0
            note = ''
            is_wknd = date_obj.weekday() in [5, 6]
            
            if date_str in presenze_user:
                presenza = presenze_user[date_str]
                tipo = presenza.get('tipo')
                if tipo == 'presenza':
                    ore_ordinarie = presenza.get('ore_lavorate', 8)
                elif tipo == 'ferie':
                    ferie = presenza.get('ore_giustificativo', 8)
                elif tipo == 'rol':
                    rol = presenza.get('ore_giustificativo', 8)
                elif tipo == 'malattia':
                    malattia = presenza.get('ore_giustificativo', 8)
                elif tipo == 'permesso':
                    ore_ordinarie = presenza.get('ore_lavorate', 0)
                    permessi = presenza.get('ore_giustificativo', 0)
                note = presenza.get('note', '')
            
            if not is_wknd:
                totali['ore_ordinarie'] += ore_ordinarie
                totali['ferie'] += ferie
                totali['rol'] += rol
                totali['permessi'] += permessi
                totali['malattia'] += malattia
            
            ws.cell(row=row, column=1).value = date_obj.strftime('%d/%m/%Y')
            ws.cell(row=row, column=2).value = day_name
            ws.cell(row=row, column=3).value = ore_ordinarie if ore_ordinarie > 0 else ''
            ws.cell(row=row, column=4).value = ferie if ferie > 0 else ''
            ws.cell(row=row, column=5).value = rol if rol > 0 else ''
            ws.cell(row=row, column=6).value = permessi if permessi > 0 else ''
            ws.cell(row=row, column=7).value = malattia if malattia > 0 else ''
            ws.cell(row=row, column=8).value = note
            
            for col in range(1, 9):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center' if col <= 7 else 'left')
                if is_wknd:
                    cell.fill = weekend_fill
            row += 1
        
        row += 1
        ws.cell(row=row, column=1).value = "TOTALI"
        ws.cell(row=row, column=1).font = Font(bold=True)
        for col_idx, key in enumerate(['ore_ordinarie', 'ferie', 'rol', 'permessi', 'malattia'], 3):
            ws.cell(row=row, column=col_idx).value = totali[key]
            ws.cell(row=row, column=col_idx).font = Font(bold=True)
        
        for col in range(1, 9):
            ws.cell(row=row, column=col).border = thin_border
            ws.cell(row=row, column=col).fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f"Presenze_{calendar.month_name[int(month)]}_{year}.xlsx")

if __name__ == '__main__':
    os.makedirs('data', exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000)
