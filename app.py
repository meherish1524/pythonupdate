
# A very simple Flask Hello World app for you to get started with...
from flask import Flask
from flask import request,render_template,redirect,url_for,send_file
import os
from docx2pdf import convert
from subprocess import  Popen
LIBRE_OFFICE = r"pythonupdate/soffice.exe"
import subprocess

#import win32com.client
#import pythoncom
#pythoncom.CoInitialize()

UPLOADER_FOLDER=''
AllOWED_EXTENSIONS={'docx'}

app = Flask(__name__)
app.config['UPLOADER_FOLDER']=UPLOADER_FOLDER

@app.route('/')
@app.route('/index',methods=['GET','POST'])
def index():
    if request.method == "POST":
        #pythoncom.CoInitialize()

        file=request.files['filename']
        if file.filename !='':
            file.save(os.path.join(app.config['UPLOADER_FOLDER'],file.filename))
        print(file.filename)
        #convert(file.filename)
        #convert(file.filename,"document.pdf")
        #return redirect('/pdf')
        #return send_file(file.filename, as_attachment=True)
        wdFormatPDF = 17
        cmd = 'libreoffice --convert-to pdf'.split() + [file.filename]
        p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
               r"", file.filename])
        print([LIBRE_OFFICE, '--convert-to', 'pdf', file.filename])
        p.communicate()



        #inputFile = os.path.abspath(r"C:\Users\mvirati\Downloads\fie.docx")
        #outputFile = os.path.abspath(r"document.pdf")
        #word = win32com.client.Dispatch('Word.Application')
        #doc = word.Documents.Open(file.filename)
        #doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
        #doc.Close()
        #word.Quit()
        return render_template("pdf.html")
        #return send_file("fie.pdf", as_attachment=True)

    return render_template("index.html")

@app.route('/pdf',methods=['GET','POST'])
def pdf():
    if request.method =="GET":
        #return send_file(file.filename,as_attachment=True)
       return send_file("document.pdf",as_attachment=True)
    print('wrong')
if __name__ == "__main__":
    app.debug=True
    app.run()
