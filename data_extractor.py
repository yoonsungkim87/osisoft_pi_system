import pythoncom
import win32com.client as win32
import pywintypes
import numpy as np
import time

server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)

point = []
iter_cnt = 1
err_cnt= 0
reason = set()

#tag = np.loadtxt('./tag.csv', dtype=np.str, delimiter=',')
tag = [
    'SE851HIS', #Turbine Speed
    'FM85404', #Starting Gas Valve
    'FM85403', #Main Gas Valve
    'FT85955', #WaterInj Flow
    'FM85310', #IGV
    'PT85178S', #P2C
    'JT86001S', #MW
    'TE8AVBP', #BP
    'TE8AVTX' #Exh
]

for x in tag:
    point.append(server.PIPoints(x).Data)
trends = []
n_samples = 250000
space = 10
unit = 's'
end_time = '2014-04-01 00:00'
#trends.append(np.linspace(space,n_samples*space,n_samples))

for p in point:
    if p is not None:
        data2 = pisdk.IPIData2(p)
        print('Extracting Data...')
        while True:
            try:
                results = data2.InterpolatedValues2(end_time+'-'+str(n_samples*space)+unit,end_time,str(space)+unit,asynchStatus=None)
                print('Successful!')
                break
            except pywintypes.com_error:
                print('Due to error, retrying...')
                pass
        tmpValue =[]
        for v in results:
            try:
                s = str(v.Value)
                tmpValue.append(float(s))
            except ValueError:
                if s == 'OFF':
                    tmpValue.append(0.0)
                elif s == 'ON':
                    tmpValue.append(1.0)
                else:
                    try:
                        tmpValue.append(tmpValue[-1])
                    except IndexError:
                        tmpValue.append(0.0)
                    finally:
                        err_cnt += 1
                        reason.add(str(v.Value))
        tmpValue.pop()
        trends.append(tmpValue)
        
print('Total Error Counter: {}'.format(err_cnt))
print('The reasons are: {}'.format(reason))
trends = np.array(trends, dtype=np.float32).transpose()

np.savetxt('result.csv', trends, delimiter=',')