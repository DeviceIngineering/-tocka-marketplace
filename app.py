import os, time, threading
from flask import Flask, request, redirect, url_for, render_template_string, send_file, flash
from processor import process_file, progress

app = Flask(__name__)
app.secret_key = 'super_secret_key'

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or file.filename == '':
            flash('Файл не выбран')
            return redirect(request.url)

        session_id = str(time.time())
        input_path = os.path.join(UPLOAD_FOLDER, f"{session_id}_{file.filename}")
        output_filename = f"result_{session_id}.xlsx"
        output_path = os.path.join(RESULT_FOLDER, output_filename)

        file.save(input_path)
        threading.Thread(
            target=process_file,
            args=(input_path, output_path, session_id),
            daemon=True
        ).start()

        return redirect(url_for('processing', session_id=session_id, filename=output_filename))

    return render_template_string('''
    <!doctype html>
    <title>Загрузка файла</title>
    <h1>Загрузите Excel-файл</h1>
    {% with msgs = get_flashed_messages() %}
      {% if msgs %}
        <ul>{% for m in msgs %}<li style="color:red;">{{ m }}</li>{% endfor %}</ul>
      {% endif %}
    {% endwith %}
    <form method="post" enctype="multipart/form-data">
      <input type="file" name="file" accept=".xlsx,.xls" required>
      <input type="submit" value="Загрузить">
    </form>
    ''')

@app.route('/processing/<session_id>/<filename>')
def processing(session_id, filename):
    return render_template_string(f'''
    <!doctype html>
    <title>Обработка файла</title>
    <h1>Статус обработки</h1>
    <p id="status">Запуск...</p>
    <script>
      function check() {{
        fetch('/status/{session_id}')
          .then(r => r.json())
          .then(d => {{
            document.getElementById('status').innerText = d.status;
            if (d.status.includes("завершена")) {{
              window.location = '/download/{filename}';
            }} else {{
              setTimeout(check, 2000);
            }}
          }});
      }}
      check();
    </script>
    ''')

@app.route('/status/<session_id>')
def status_api(session_id):
    return {"status": progress.get(session_id, "Нет данных")}

@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(RESULT_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    flash("Файл не найден")
    return redirect(url_for('upload_file'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)
