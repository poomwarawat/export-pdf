from docxtpl import DocxTemplate
from docx2pdf import convert

from flask import Flask, send_file, request, after_this_request, Response, redirect, render_template
from flask_cors import CORS
import uuid
import os
from os import listdir
from os.path import isfile, join
import json
from werkzeug.utils import secure_filename
from fontTools.ttLib import TTFont

app = Flask(__name__)

CORS(app)

UPLOAD_FOLDER = "./template"
UPLOAD_FONT_FOLDER = "./fonts"
ALLOWED_EXTENSIONS = {'docx', 'ttf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['UPLOAD_FONT_FOLDER'] = UPLOAD_FONT_FOLDER


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/font', methods=['GET'])
def getAllfont():
    tmp = os.popen("fc-list").read()
    font_list = tmp.splitlines()

    font = {
        'font': font_list,
        'font_size': len(font_list)
    }

    return Response(json.dumps(font),  mimetype='application/json')


def install_font(filename):
    font = TTFont(f"fonts/{filename}")
    font.save(f"../../../Library/Fonts/{filename}")


@app.route('/font', methods=['POST'])
def upload_font():
    print('-> start upload font')
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file to upload"
        file = request.files['file']
        if file.filename == '':
            return "Filename is empty"
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FONT_FOLDER'], filename))
            install_font(filename)
            return "Success"


@app.route('/pdf/template', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file to upload"
        file = request.files['file']
        if file.filename == '':
            return "Filename is empty"
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return "Success"


@app.route('/pdf/template', methods=['GET'])
def get_templates():
    relevant_path = "./template"
    included_extensions = ['docx']
    file_names = [fn for fn in os.listdir(relevant_path)
                  if any(fn.endswith(ext) for ext in included_extensions)]
    file = {
        "file": file_names
    }

    return Response(json.dumps(file),  mimetype='application/json')


@app.route('/pdf/template/<path:filename>', methods=['GET'])
def downloadFile(filename):
    # For windows you need to use drive name [ex: F:/Example.pdf]
    path = f"./template/{filename}"
    return send_file(path, as_attachment=True)


@app.route('/pdf/template/<path:filename>', methods=['DELETE'])
def deleteFile(filename):
    os.remove(f"template/{filename}")
    return "Success"


@app.route('/pdf/export', methods=['POST'])
def export_pdf():
    body = request.json
    template_name = body['templateName']
    print(template_name)
    context = body['context']
    id = uuid.uuid4()
    filename = template_name.split(".")[0]

    tpl = DocxTemplate(f"template/{template_name}")
    tpl.render(context)

    tpl.save(f"output/{filename}.docx")
    convert(f"output/{filename}.docx", f"pdf/{filename}.pdf")

    return send_file(f"pdf/{filename}.pdf", mimetype="application/application/pdf")


if __name__ == '__main__':
    app.run()
