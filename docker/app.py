from flask import request,Flask,jsonify
import os
import subprocess

app = Flask(__name__)

@app.route('/status',  methods=['GET'])
def status():
  res={}
  try:
    res['code']=200
    res['status']=True
  except Exception as e:
    res['code']=500
    res['status']=False
    res['errMsg']=str(e)
  return jsonify(res)
@app.route('/',  methods=['POST','GET'])
def index():
  res={}
  res['status']=False
  res['code']=200
  res['errMsg']=""
  try:
    # data = request.args
    # print(data['text'].encode('utf-8').decode('unicode_escape'))
    err=''
    if(request.method=='GET'):
      data = request.args
    elif(request.method=='POST'):
      data = request.get_json()
    if(not data):
      res['code']=40001
      res['errMsg']='request data is null'
    else:
      infile=data.get('infile')
      extension=data.get('extension')
      outdir=data.get('outdir')
      outfile=data.get('outfile')
      if(not infile):
        res['code']=40002
        res['errMsg']='infile is null'
      elif(not extension):
        res['code']=40003
        res['errMsg']='extension is null'
      elif(not outdir):
        res['code']=40004
        res['errMsg']='outdir is null'
      else:
        result = subprocess.run(["libreoffice", "--headless", "--convert-to", extension,infile,"--outdir", outdir], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        res['data']={"infile":infile, "extension":extension, "outdir": outdir}
        res['stdout']=result.stdout
        if(result.returncode != 0):
          res['code']=500
          res['errMsg']=str(result.stderr)
    if(res['errMsg']!=""):
      pass
    else:
      res['code']=200
      res['status']=True
  except Exception as e:
    res['code']=500
    res['errMsg']=str(e)
  return jsonify(res)

if __name__ == '__main__'
  app.debug = os.environ.get('FLASK_DEBUG') == "True" # 设置调试模式，生产模式的时候要关掉debug
  app.run(port=5055, host='0.0.0.0') # Container运行必须: host='0.0.0.0'