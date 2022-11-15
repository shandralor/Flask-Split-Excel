import pandas as pd
import openpyxl
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename
import os

#VARIABLES
PATH = ''
ALLOWED_EXTENSIONS = set(['xlsx'])

#FLASK APP
app = Flask(__name__)

#ALLOWED FILES FUNCTION
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower()in ALLOWED_EXTENSIONS

#PANDAS EXCEL FUNCTION
def split_excel(PATH):
    df=pd.read_excel(PATH,sheet_name=None)
    for p in df:
        sheet=(df[p])
        sheet.to_excel(f"downloads/{p}.xlsx", index=False)

def clear_downloads():
    path = 'downloads'
    for file_to_remove in os.listdir(path):
        if os.path.isfile(os.path.join(path, file_to_remove)):
            os.remove(f"downloads/"+ file_to_remove)
 
           
#ROUTES
@app.route('/', methods=["GET", "POST"])
def index():
    if request.method=="POST":
        clear_downloads()
        excel_file = request.files['file']
        if excel_file and allowed_file(excel_file.filename):
            filename = secure_filename(excel_file.filename)
            saved_loc = os.path.join('uploads', filename)
            excel_file.save(saved_loc)
            split_excel(saved_loc)
            os.remove(saved_loc)
        return redirect(url_for("download"))
        
        
    return render_template("index.html")

@app.route('/download')
def download():
    return render_template("download.html", files=os.listdir('downloads'))

@app.route('/download/<excel>')
def download_excel(excel):
    return send_from_directory('downloads', excel)


#RUN MAIN FUNCTION
if __name__=='__main__':
    app.run(debug=True)