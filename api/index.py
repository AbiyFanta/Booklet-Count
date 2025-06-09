from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os
import tempfile

app = Flask(__name__)

HTML_FORM = """
<!DOCTYPE html>
<html>
<head><title>Excel Upload</title></head>
<body>
    <h2>Upload an Excel File (.xlsx)</h2>
    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx" required>
        <button type="submit">Generate Report</button>
    </form>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        if file and file.filename.endswith(".xlsx"):
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

    return render_template_string(HTML_FORM)

# Vercel entrypoint
def handler(request, context):
    return app(request, context)
