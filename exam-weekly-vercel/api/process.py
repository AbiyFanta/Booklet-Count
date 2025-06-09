# api/process.py

from flask import Flask, request, send_file
import pandas as pd
import os
import tempfile

app = Flask(__name__)

@app.route("/", methods=["POST"])
def handle_upload():
    file = request.files.get("file")
    if not file or not file.filename.endswith(".xlsx"):
        return "Invalid file", 400

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file.filename)
        output_path = os.path.join(tmpdir, "exam_weekly.xlsx")
        file.save(input_path)

        df = pd.read_excel(input_path)
        df['EXAM DATE'] = pd.to_datetime(df['EXAM DATE'])
        df['Week'] = df['EXAM DATE'].dt.to_period('W').apply(lambda r: r.start_time)
        pivot_df = df.pivot_table(index='Exam Site', columns='Week', aggfunc='size', fill_value=0)
        pivot_df = pivot_df.sort_index(axis=1)
        pivot_df.to_excel(output_path)

        return send_file(output_path, as_attachment=True)

def handler(request, context):  # for Vercel
    return app(request, context)
