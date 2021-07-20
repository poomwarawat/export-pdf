from docxtpl import DocxTemplate
from docx2pdf import convert

from flask import Flask, send_file, request, after_this_request, Response, redirect, render_template
from flask_cors import CORS
import uuid
import os
from os import listdir
from os.path import isfile, join
import json

app = Flask(__name__)
CORS(app)


@app.route('/pdf/template', methods=['GET'])
def get_templates():
    relevant_path = "./template"
    included_extensions = ['docx']
    file_names = [fn for fn in os.listdir(relevant_path)
                  if any(fn.endswith(ext) for ext in included_extensions)]
    file = {
        "file": file_names
    }

    print(file)
    return Response(json.dumps(file),  mimetype='application/json')


@app.route('/pdf/template/<path:filename>', methods=['GET'])
def downloadFile(filename):
    # For windows you need to use drive name [ex: F:/Example.pdf]
    path = f"./template/{filename}"
    return send_file(path, as_attachment=True)


@app.route('/pdf/export', methods=['POST'])
def export_pdf():
    body = request.json
    template_name = body['templateName']
    context = body['context']
    id = uuid.uuid4()
    print(template_name)
    print(context)

    tpl = DocxTemplate(f"template/{template_name}")
    tpl.render(context)

    tpl.save(f"output/{id}.docx")
    convert(f"output/{id}.docx", f"pdf/{id}.pdf")

    @after_this_request
    def remove_file(response):
        try:
            os.remove(f"./output/{id}.docx")
            os.remove(f"./output/{id}.pdf")
        except Exception as error:
            app.logger.error(
                "Error removing or closing downloaded file handle", error)

        return response

    return send_file(f"./pdf/{id}.docx", mimetype="application/application/docx")


if __name__ == '__main__':
    app.run()

# context = {
#     'projectName': 'โครงการจัดซื้ออุปกรณ์สร้างโรงงาน',
#     'contractNumber': '86543567',
#     'leaderName': 'มายวัน',
#     'workDetails': 'จัดซ์้อวัสดุและอุปกรณ์และก่อสร้างตั้งแต่ 18 มกราคม - 30 ธันวาคม',
#     'tbl_contents': [
#         {'name': 'วรวัชร', 'lastname': 'ไล้เลิศ'},
#         {'name': 'พิชัย', 'lastname': 'มาแว้ว', },
#         {'name': 'โทริโก้', 'lastname': 'นั้นโก้จริงๆ'},
#         {'name': 'ยูด้า', 'lastname': 'ดายู้'},
#     ],
#     'currentDate': now.strftime("%d/%m/%Y, %H:%M:%S")
# }
