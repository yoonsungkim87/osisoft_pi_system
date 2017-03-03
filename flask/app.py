from flask import Flask
from flask import Markup
from flask import Flask
from flask import render_template
app = Flask(__name__)
 
@app.route("/")
def chart():
    import pythoncom
    import win32com.client as win32
    import numpy as np

    server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
    pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
    
    tag, point = [None] * 2, [None] * 2
    tag[0] = '70MKG31CQ001A_XQ01'
    tag[1] = '80MKG31CQ001A_XQ01'
    point[0] = server.PIPoints(tag[0]).Data
    point[1] = server.PIPoints(tag[1]).Data
    trends = []
    n_samples = 1000
    for p in point:
        data2 = pisdk.IPIData2(p)
        results = data2.InterpolatedValues2('*-'+str(n_samples)+'h','*','1h',asynchStatus=None)
        tmpValue =[]
        tmpTime = []
        i = 1 - n_samples
        for v in results:
            try:
                s = str(v.Value)
                tmpValue.append(float(s))
                tmpTime.append(i)
            except ValueError:
                pass
            i = i + 1
        tmpValue.pop()
        tmpTime.pop()
        trends.append([tmpTime, tmpValue])
    labels = trends[0][0]
    value1 = trends[0][1]
    value2 = trends[1][1]
    
    return render_template('chart.html', labels = labels, value1 = value1, value2 = value2)
 
if __name__ == "__main__":
    app.run(host='172.18.104.115')