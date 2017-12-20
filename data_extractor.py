import pythoncom
import win32com.client as win32
import numpy as np
import time

server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)

point = []
iter_cnt = 1
err_cnt= 0
reason = set()

#tag = np.loadtxt('./tag.csv', dtype=np.str, delimiter=',')
tag = ['51HBK27CP101XQ50',
       '51HBK27CT101XQ50',
       'AIT-512-501B',
       'OCB02E7005-OUT',
       'TIT-512-504',
       'TIT-512-503A',
       'TIT-512-503B',
       'TIT-512-501',
       'PIT-512-501',
       'FIT-512-502-CAL',
       'HV-512-504',
       'PIT-512-552',
       'ZT-FCV-512-501',
       'OCB02E7008-OUT',
       '51HNE10CQ102']

for x in tag:
    point.append(server.PIPoints(x).Data)
    trends = []
    n_samples = 1000000
    space = 10
    unit = 's'
    end_time = '2017-12-19 00:00'
    #trends.append(np.linspace(space,n_samples*space,n_samples))

    for p in point:
        data2 = pisdk.IPIData2(p)
        print('Extracting Data...')
        results = data2.InterpolatedValues2(end_time+'-'+str(n_samples*space)+unit,end_time,str(space)+unit,asynchStatus=None)
        time.sleep(1)
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