from flask import Flask
from flask import Markup
from flask import render_template
app = Flask(__name__)
 
@app.route("/")
def chart():
    import pythoncom
    import win32com.client as win32
    import numpy as np

    server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
    pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
    
    point = []
    tag = ['70MKG31CQ001A_XQ01','80MKG31CQ001A_XQ01']
    for x in tag:
        point.append(server.PIPoints(x).Data)
    trends = []
    n_samples = 1000
    space = 1
    unit = 'h'
    
    for p in point:
        data2 = pisdk.IPIData2(p)
        results = data2.InterpolatedValues2('*-'+str(n_samples)+unit,'*',str(space)+unit,asynchStatus=None)
        tmpValue =[]
        tmpTime = []
        i = 1 - space * n_samples
        for v in results:
            try:
                s = str(v.Value)
                tmpValue.append(float(s))
                tmpTime.append(i)
            except ValueError:
                pass
            i += space
        tmpValue.pop()
        tmpTime.pop()
        trends.append([tmpTime, tmpValue])
        
    labels = trends[0][0]
    value1 = trends[0][1]
    value2 = trends[1][1]
    
    return render_template('chart.html', labels = labels, value1 = value1, value2 = value2)
 
if __name__ == "__main__":
    app.run(host='me')