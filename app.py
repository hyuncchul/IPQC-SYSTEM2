from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file, flash
from datetime import date, datetime
import json, os
from database import init_db, save_entry, get_entry, get_daily_status, save_handover, save_abnormality, get_history_dates, get_all_entries_for_date
from machine_config import MACHINE_LIST, MACHINE_CONFIG
from export_excel import export_daily_excel

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'qc_ipqc_secret_2025')
app.jinja_env.globals.update(enumerate=enumerate)

init_db()

SHIFT_LABELS = {'morning': '早班', 'night': '夜班', 'first_piece': '首件'}

@app.route('/')
def index():
    today = date.today().strftime('%Y-%m-%d')
    return redirect(url_for('dashboard', date_str=today))

@app.route('/dashboard/<date_str>')
def dashboard(date_str):
    data = get_daily_status(date_str)
    entered = {(e['machine_id'], e['shift']) for e in data['entries']}
    handover_map = {h['machine_id']: h for h in data['handovers']}
    machine_status = {}
    for m in MACHINE_LIST:
        machine_status[m] = {
            'morning': ('morning' in [e['shift'] for e in data['entries'] if e['machine_id']==m]),
            'night': ('night' in [e['shift'] for e in data['entries'] if e['machine_id']==m]),
            'first_piece': ('first_piece' in [e['shift'] for e in data['entries'] if e['machine_id']==m]),
            'handover': m in handover_map,
            'handover_data': handover_map.get(m, {}),
        }
    today = date.today().strftime('%Y-%m-%d')
    return render_template('dashboard.html',
        date_str=date_str,
        today=today,
        machine_list=MACHINE_LIST,
        machine_status=machine_status,
        abnormalities=data['abnormalities'],
        shift_labels=SHIFT_LABELS)

@app.route('/entry/select')
def entry_select():
    today = date.today().strftime('%Y-%m-%d')
    return render_template('entry_select.html',
        machine_list=MACHINE_LIST,
        today=today,
        shifts=SHIFT_LABELS)

@app.route('/entry/<machine_id>/<shift>', methods=['GET'])
def entry_form(machine_id, shift):
    if machine_id not in MACHINE_CONFIG:
        flash('잘못된 기계번호입니다.', 'danger')
        return redirect(url_for('entry_select'))
    today = date.today().strftime('%Y-%m-%d')
    date_str = request.args.get('date', today)
    config = MACHINE_CONFIG[machine_id]
    existing = get_entry(date_str, machine_id, shift)
    return render_template('entry_form.html',
        machine_id=machine_id,
        shift=shift,
        shift_label=SHIFT_LABELS.get(shift, shift),
        date_str=date_str,
        config=config,
        existing=existing,
        shift_labels=SHIFT_LABELS)

@app.route('/entry/<machine_id>/<shift>', methods=['POST'])
def entry_submit(machine_id, shift):
    form = request.form
    date_str = form.get('date', date.today().strftime('%Y-%m-%d'))
    part_no = form.get('part_no','')
    lot_no = form.get('lot_no','')
    submitted_by = form.get('submitted_by','')
    notes = form.get('notes','')

    config = MACHINE_CONFIG.get(machine_id, {'visual':[],'eol':[],'dims':[]})

    visual_items = []
    for i, name in enumerate(config['visual']):
        visual_items.append({
            'name': name,
            'result': form.get(f'visual_result_{i}',''),
            'rejected_lot': form.get(f'visual_lot_{i}','')
        })

    eol_items = []
    for i, name in enumerate(config['eol']):
        eol_items.append({
            'name': name,
            'result': form.get(f'eol_result_{i}',''),
            'rejected_lot': form.get(f'eol_lot_{i}','')
        })

    dim_items = []
    for i, name in enumerate(config['dims']):
        dim_items.append({
            'name': name,
            'result': form.get(f'dim_result_{i}','')
        })

    try:
        save_entry(date_str, machine_id, shift, part_no, lot_no, submitted_by, notes,
                   visual_items, eol_items, dim_items)
        flash(f'{machine_id} {SHIFT_LABELS.get(shift,shift)} 데이터가 저장되었습니다!', 'success')
    except Exception as e:
        flash(f'저장 실패: {str(e)}', 'danger')

    return redirect(url_for('dashboard', date_str=date_str))

@app.route('/handover', methods=['GET','POST'])
def handover():
    today = date.today().strftime('%Y-%m-%d')
    if request.method == 'POST':
        date_str = request.form.get('date', today)
        for m in MACHINE_LIST:
            last_batch = request.form.get(f'last_batch_{m}','')
            reason = request.form.get(f'reason_{m}','')
            if last_batch or reason:
                save_handover(date_str, m, last_batch, reason)
        flash('停機&尾批交接 데이터가 저장되었습니다!', 'success')
        return redirect(url_for('dashboard', date_str=date_str))

    date_str = request.args.get('date', today)
    data = get_daily_status(date_str)
    handover_map = {h['machine_id']: h for h in data['handovers']}
    return render_template('handover.html',
        machine_list=MACHINE_LIST,
        date_str=date_str,
        today=today,
        handover_map=handover_map)

@app.route('/abnormality', methods=['GET','POST'])
def abnormality():
    today = date.today().strftime('%Y-%m-%d')
    if request.method == 'POST':
        date_str = request.form.get('date', today)
        machine_id = request.form.get('machine_id','')
        shift = request.form.get('shift','')
        description = request.form.get('description','')
        cause = request.form.get('cause','')
        countermeasure = request.form.get('countermeasure','')
        if machine_id and description:
            save_abnormality(date_str, machine_id, shift, description, cause, countermeasure)
            flash('이상 발생 내용이 저장되었습니다!', 'success')
        return redirect(url_for('dashboard', date_str=date_str))

    date_str = request.args.get('date', today)
    data = get_daily_status(date_str)
    return render_template('abnormality.html',
        machine_list=MACHINE_LIST,
        date_str=date_str,
        today=today,
        abnormalities=data['abnormalities'],
        shifts=SHIFT_LABELS)

@app.route('/history')
def history():
    dates = get_history_dates(60)
    return render_template('history.html', dates=dates)

@app.route('/view/<date_str>')
def view_date(date_str):
    all_data = get_all_entries_for_date(date_str)
    return render_template('view_date.html',
        date_str=date_str,
        all_data=all_data,
        shift_labels=SHIFT_LABELS,
        machine_list=MACHINE_LIST)

@app.route('/export/<date_str>')
def export(date_str):
    all_data = get_all_entries_for_date(date_str)
    buf = export_daily_excel(date_str, all_data)
    filename = f"QC_{date_str.replace('-','')}.xlsx"
    return send_file(buf, download_name=filename,
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/api/status/<date_str>')
def api_status(date_str):
    data = get_daily_status(date_str)
    return jsonify(data)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
