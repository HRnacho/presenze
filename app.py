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
    from openpyxl.styles import PatternFill, Alignment
    
    # Carica il template Excel
    template_path = os.path.join('templates', 'Foglio_presenze_UDINE_Dicembre_2025.xlsx')
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    # Aggiorna il periodo nelle celle G-L unite (riga 1)
    month_names = ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno',
                   'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre']
    
    # Trova e unmerge le celle G1:L1 se sono merged
    merged_cells_to_unmerge = []
    for merged_cell_range in ws.merged_cells.ranges:
        # Controlla se G1 è nel range merged
        if 'G1' in merged_cell_range:
            merged_cells_to_unmerge.append(merged_cell_range)
    
    for merged_range in merged_cells_to_unmerge:
        ws.unmerge_cells(str(merged_range))
    
    # Scrivi il mese in G1
    ws['G1'] = f"{month_names[int(month)-1]}-{str(year)[-2:]}"
    
    # Ri-mergia le celle G1:L1
    ws.merge_cells('G1:L1')
    ws['G1'].alignment = Alignment(horizontal='center', vertical='center')
    
    # Carica i dati delle presenze
    data = load_data()
    key = f"{year}-{month.zfill(2)}"
    presenze_mese = data.get(key, {})
    
    num_days = calendar.monthrange(int(year), int(month))[1]
    
    # Mappa username -> riga di partenza nel template (riga Ord.)
    user_rows = {
        'gianluca': 4,   # Bittoni Gianluca inizia alla riga 4
        'ignacio': 8,    # Sorcaburu Ciglieri Ignacio inizia alla riga 8
        'simone': 14,    # Mascellari Simone inizia alla riga 14
    }
    
    # Festivi italiani 2025-2026
    holidays = {
        1: [1, 6],           # Capodanno, Epifania
        2: [],               # Febbraio 2026
        3: [],               # Marzo 2026
        4: [20, 21],         # Pasqua, Pasquetta (2025)
        5: [1],              # Festa del Lavoro
        6: [2],              # Festa della Repubblica
        7: [],               # Luglio
        8: [15],             # Ferragosto
        9: [],               # Settembre
        10: [],              # Ottobre
        11: [1],             # Ognissanti
        12: [8, 25, 26]      # Immacolata, Natale, Santo Stefano
    }
    
    current_month_holidays = holidays.get(int(month), [])
    
    # Popola i dati per ogni dipendente
    for username, base_row in user_rows.items():
        if username not in presenze_mese:
            continue
            
        presenze_user = presenze_mese[username]
        
        # Per ogni giorno del mese
        for day in range(1, num_days + 1):
            # Calcola l'indice di colonna (D=4, E=5, ... colonna 3 + giorno)
            col_index = 3 + day  # D è la colonna 4, quindi 3+1=4
            col_letter = openpyxl.utils.get_column_letter(col_index)
            
            date_str = f"{year}-{month.zfill(2)}-{str(day).zfill(2)}"
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            
            is_weekend = date_obj.weekday() in [5, 6]  # Sabato=5, Domenica=6
            is_holiday = day in current_month_holidays
            
            # Colora weekend e festivi in rosso e lascia vuoto
            if is_weekend or is_holiday:
                red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                for offset in range(4):  # 4 righe: Ord, Str, Ass, Giust
                    cell = ws[f'{col_letter}{base_row + offset}']
                    cell.fill = red_fill
                    cell.value = None
                continue
            
            # Se ci sono dati per questo giorno
            if date_str in presenze_user:
                presenza = presenze_user[date_str]
                tipo = presenza.get('tipo', 'presenza')
                ore_lavorate = presenza.get('ore_lavorate', 0)
                ore_assenza = presenza.get('ore_assenza', 0)
                
                # Calcola ore ordinarie e straordinari
                ore_ordinarie = max(0, min(8, ore_lavorate) - ore_assenza)
                ore_straordinari = max(0, ore_lavorate - 8) if ore_lavorate > 8 else 0
                
                # Riga Ord. (ore ordinarie)
                cell_ord = ws[f'{col_letter}{base_row}']
                if ore_ordinarie > 0:
                    cell_ord.value = f"{ore_ordinarie:.2f}".replace('.', ',')
                    cell_ord.alignment = Alignment(horizontal='center')
                else:
                    cell_ord.value = None
                
                # Riga Str. (straordinari)
                cell_str = ws[f'{col_letter}{base_row + 1}']
                if ore_straordinari > 0:
                    cell_str.value = f"{ore_straordinari:.2f}".replace('.', ',')
                    cell_str.alignment = Alignment(horizontal='center')
                else:
                    cell_str.value = None
                
                # Riga Ass. (assenze)
                cell_ass = ws[f'{col_letter}{base_row + 2}']
                if ore_assenza > 0:
                    cell_ass.value = f"{ore_assenza:.2f}".replace('.', ',')
                    cell_ass.alignment = Alignment(horizontal='center')
                else:
                    cell_ass.value = None
                
                # Riga Giust. (giustificativo)
                cell_giust = ws[f'{col_letter}{base_row + 3}']
                if tipo != 'presenza' and ore_assenza > 0:
                    codice_map = {
                        'ferie': 'FERIE',
                        'rol': 'ROL',
                        'malattia': 'MALATTIA',
                        'permesso': 'PERMESSO'
                    }
                    cell_giust.value = codice_map.get(tipo, '')
                    cell_giust.alignment = Alignment(horizontal='center')
                else:
                    cell_giust.value = None
    
    # Salva in BytesIO per il download
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"Foglio_presenze_UDINE_{month_names[int(month)-1].title()}_{year}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    os.makedirs('data', exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000)
