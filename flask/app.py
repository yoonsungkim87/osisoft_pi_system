from flask import Flask, Markup, render_template, request

app = Flask(__name__)
 
@app.route("/<path:varargs>", methods=['GET'])
def chart(varargs=None):
    import pythoncom
    import win32com.client as win32
    import numpy as np

    server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
    pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
    
    point = []
    tags= varargs.split("/")
    print(tags)
    for x in tags:
        point.append(server.PIPoints(x).Data)
    trends = []
    n_samples = 100
    space = 1
    unit = 'h'
    
    init = True
    for p in point:
        data2 = pisdk.IPIData2(p)
        results = data2.InterpolatedValues2('*-'+str(n_samples)+unit,'*',str(space)+unit,asynchStatus=None)
        tmpValue =[]
        tmpTime = []
        i = 0
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
        if init:
            trends.append(tmpTime)
            init = False
        trends.append(tmpValue)
    trends = np.array(trends).transpose().tolist()
    return render_template('googlechart.html', tag=tags, trend=trends)
 
if __name__ == "__main__":
    app.run(host='me')