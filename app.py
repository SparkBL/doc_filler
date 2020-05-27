from flask import Flask, flash, request, redirect, url_for,render_template,send_from_directory
from werkzeug.utils import secure_filename
import templater as te
import os,glob,zipfile
import uuid
UPLOAD_FOLDER = 'uploads/'
RESULT_FOLDER = 'docs/'
RESULT_ZIP = ''
ALLOWED_DOC = set(['doc','docx'])
ALLOWED_TABLE = set(['xlsx'])

app = Flask(__name__)
app.config['UPLOAD_FOLDER']=UPLOAD_FOLDER
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/=='


@app.route("/",methods = ['GET','POST'])
def home():
    message =""
    if request.method == 'POST':
        template = request.files['template']
        data = request.files['data']
        if not data or not template or not allowed_file(template.filename,ALLOWED_DOC) or not allowed_file(data.filename,ALLOWED_TABLE):
            flash("Неверный формат или не выбраны файлы")
            return redirect(request.url)
        template.save(os.path.join(app.config['UPLOAD_FOLDER'],secure_filename(template.filename)))
        data.save(os.path.join(app.config['UPLOAD_FOLDER'],secure_filename(data.filename)))
        clear(RESULT_FOLDER)
        RESULT_ZIP = str(uuid.uuid4()) + '.zip'
        te.fill_docs(os.path.join(app.config['UPLOAD_FOLDER'],secure_filename(data.filename)),os.path.join(app.config['UPLOAD_FOLDER'],secure_filename(template.filename)),RESULT_FOLDER)
        zipdir(RESULT_FOLDER,zipfile.ZipFile(UPLOAD_FOLDER +RESULT_ZIP, 'w',zipfile.ZIP_DEFLATED))
        return render_template("home.html",filename=RESULT_ZIP)
    if request.method == 'GET':
        
        clear(UPLOAD_FOLDER)
        return render_template("home.html")

@app.route("/download/<filename>")
def download(filename):
            return send_from_directory(directory=UPLOAD_FOLDER,filename=filename,as_attachment=True,attachment_filename=filename)




def allowed_file(filename,allowed):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in allowed

def clear(path):
    files = glob.glob(path + "*")
    for f in files:
        os.remove(f)


def zipdir(path,ziph):
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(path, file))

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.mkdir(UPLOAD_FOLDER)
    app.run()