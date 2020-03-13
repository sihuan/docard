from flask import Flask, render_template, request, send_file
import work

from config import SKEY, SVALUE, PORT, DEBUG

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/check')
def check():
    return render_template('check.html')


@app.route('/checkdata', methods=['POST'])
def checkdata():
    data = request.json
    cr = data['classroom']
    s = work.checkdata(cr)
    return s


@app.route('/checkall')
def checkall():
    return render_template('checkall.html')


@app.route('/checkalldata', methods=['POST'])
def checkalldata():
    s = work.checkalldata()
    return s


@app.route('/tips', methods=['POST'])
def tips():
    data = request.json
    if data[SKEY] == SVALUE:
        try:
            work.tipall()
            return{
                'status': 1,
            }
        except Exception as e:
            return {
                'status': 0,
                'msg': str(e)
            }
    else:
        return{
            'status': 'mmp你想干嘛？'
        }


@app.route('/getstudent', methods=['POST'])
def getstudent():
    data = request.json
    sid = data['sid']
    print(sid)
    s = work.findstudent(sid)
    if s:
        return {
            'status': 1,
            'student': s,
        }
    else:
        return {
            'status': 0,
        }


@app.route('/docard', methods=['POST'])
def docard():
    data = request.json
    try:
        if work.doCard(data['sid'], data['now'], data['te'], data['know'], data['change']):
            return {
                'status': 1,
            }
    except:
        return {
            'status': 0,
        }


@app.route('/reload', methods=['POST'])
def reload():
    data = request.json
    if data[SKEY] == SVALUE:
        try:
            work.r.flushall()
            work.loadStudent('test.xlsx')
            return {
                'status': 1
            }
        except Exception as e:
            return {
                'status': 0,
                'msg': str(e)
            }
    else:
        return{
            'status': 'mmp你想干嘛？'
        }


@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    if filename == 'page':
        return render_template('download.html')
    work.export(filename+'.xlsx')
    return send_file(filename+'.xlsx', as_attachment=True)


if __name__ == '__main__':
    app.run(debug=DEBUG, host='0.0.0.0', port=PORT)
