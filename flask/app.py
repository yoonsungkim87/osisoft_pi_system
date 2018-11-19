from flask import Flask, Markup, render_template, request
from dateutil.parser import parse
import pythoncom, pywintypes
import win32com.client as win32

server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)


app = Flask(__name__)
 
@app.route("/trend/<string:varargs>", methods=['GET'])
def chart(varargs=None):
    
    point = []
    tags= varargs.split(",")
    print(tags)
    for x in tags:
        try:
            point.append(server.PIPoints(x).Data)
        except pywintypes.com_error as e:
            print('REASON: '+repr(e))
    trends = []
    period = '0.1h'
    
    i =0
    for p in point:
        data2 = pisdk.IPIData2(p)
        #results = data2.InterpolatedValues2('*-'+str(n_samples)+unit,'*',str(space)+unit,asynchStatus=None)
        results = data2.RecordedValues('*-'+period,'*',3,"",0,None)
        tmp =[]
        for v in results:
            try:
                s = str(v.Value)
                dt = parse(str(v.TimeStamp.LocalDate))
                time = (dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second, dt.microsecond)
                element =[]
                for j in range(i):
                    element.append(None)
                element.append(float(s))
                for k in range(len(tags) - i -1):
                    element.append(None)
                tmp.append([time, element])
            except ValueError:
                pass
        tmp.pop()
        trends.append(tmp)
        i += 1
    #print(trends)
    return render_template('trend.html', tag=tags, trend=trends)
 
if __name__ == "__main__":
    app.run(host='me')