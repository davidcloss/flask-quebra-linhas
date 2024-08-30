from flask import Flask, request, send_file, render_template
import pandas as pd
import io
import zipfile
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def convert_scientific_to_int(df, column_name):
    if column_name in df.columns:
        try:
            df[column_name] = df[column_name].astype(str).str.replace(',', '.').astype(float).astype(int)
        except ValueError:
            print(f"Não foi possível converter a coluna {column_name} para inteiro.")
    return df

def split_and_zip(df, lines_per_file, base_filename):
    df = convert_scientific_to_int(df, 'CPF/CNPJ')
    df = df.drop_duplicates(subset='CPF/CNPJ')

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        total_rows = len(df)
        num_parts = (total_rows + lines_per_file - 1) // lines_per_file

        for i in range(num_parts):
            start_row = i * lines_per_file
            end_row = min(start_row + lines_per_file, total_rows)
            part_df = df.iloc[start_row:end_row]

            csv_buffer = io.StringIO()
            part_df.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)

            zipf.writestr(f"{base_filename}_parte_{i+1}.csv", csv_buffer.getvalue())
    
    buffer.seek(0)
    return buffer

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            lines_per_file = int(request.form.get('lines_per_file', 500))

            # Lê a planilha diretamente do BytesIO
            file_bytes = io.BytesIO(file.read())
            df = pd.read_excel(file_bytes)

            # Divide e salva os arquivos no buffer ZIP
            zip_buffer = split_and_zip(df, lines_per_file, filename.rsplit('.', 1)[0])

            return send_file(zip_buffer, as_attachment=True, download_name=f"{filename.rsplit('.', 1)[0]}.zip")

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

