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
    n_samples = 100000
    space = 0.1
    unit = 'm'
    end_time = '2017-12-01 00:00'
    trends.append(np.linspace(space,n_samples*space,n_samples))

    for p in point:
        data2 = pisdk.IPIData2(p)
        results = data2.InterpolatedValues2(end_time+'-'+str(n_samples*space)+unit,end_time,str(space)+unit,asynchStatus=None)
        tmpValue =[]
        for v in results:
            try:
                s = str(v.Value)
                tmpValue.append(float(s))
            except ValueError:
                tmpValue.append(tmpValue[-1])
        tmpValue.pop()
        trends.append(tmpValue)
        
#print(np.array(trends[1]).shape)
trends = np.array(trends, dtype=np.float32).transpose()

np.savetxt('result.csv', trends, delimiter=',')

    