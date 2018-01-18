import pythoncom
import win32com.client as win32
import pywintypes
import numpy as np

server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)

point = []
iter_cnt = 1
err_cnt= 0
reason = set()

#tag = np.loadtxt('./tag.csv', dtype=np.str, delimiter=',')
tag = [
    'SE851HIS', # Turbine Speed
    'TE85147', # T1C
    'PT85169', # P1C
    'PT023904', # Fuel Gas Pressure
    'FM85404', # Starting Valve DMD
    'FM85403', # Gas Valve DMD
    'FM85310', # IGV DMD
    'FM85951', # Water Injection CV DMD
    
    'JT86001S', # Active Power
    'TE851DMD', # TEMP CNT DMD
    
    'TE85117', # DC2-1
    'TE85118', # DC2-2
    'TE85119', # DC3-1
    'TE85120', # DC3-2
    'TE85121', # DC4-1
    'TE85122', # DC4-2
    
    'TE8SPBP', # SPT BP
    'TE8AVBP', # AVG BP
    
    'TE85101S', # BP1
    'TE85102S', # BP2
    'TE85103S', # BP3
    'TE85104S', # BP4
    'TE85105S', # BP5
    'TE85106S', # BP6
    'TE85107S', # BP7
    'TE85108S', # BP8
    'TE85109S', # BP9
    'TE85110S', # BP10
    'TE85111S', # BP11
    'TE85112S', # BP12
    'TE85113S', # BP13
    'TE85114S', # BP14
    
    'TE8SPTX', # SPT EXH
    'TE8AVTX', # AVG EXH
    
    'TE85151', # EXH1
    'TE85152', # EXH2
    'TE85153', # EXH3
    'TE85154', # EXH4
    'TE85155', # EXH5
    'TE85156', # EXH6
    'TE85157', # EXH7
    'TE85158', # EXH8
    'TE85159', # EXH9
    'TE85160', # EXH10
    'TE85161', # EXH11
    'TE85162', # EXH12
    'TE85163', # EXH13
    'TE85164', # EXH14
    'TE85165', # EXH15
    'TE85166', # EXH16
    
    'PT85178S', # P2C
    'TE85315S', # T2C
    'PT75168', # Inlet DP
    
]

for x in tag:
    point.append(server.PIPoints(x).Data)
trends = []
n_samples = int(7*24*60)
space = 1
unit = 'm'
end_time = '2018-01-17 00:00'
#trends.append(np.linspace(space,n_samples*space,n_samples))

for p in point:
    if p is not None:
        data2 = pisdk.IPIData2(p)
        print('Extracting Data...')
        while True:
            try:
                results = data2.InterpolatedValues2(end_time+'-'+str(n_samples*space)+unit,end_time,str(space)+unit,asynchStatus=None)
                print('**************************Successful!')
                break
            except pywintypes.com_error:
                print('Error occured, retrying...')
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
        
print('Total Error Counter: ', end='')
print(err_cnt)

print('Reason: ', end='')
print(*reason if reason else '', sep=', ')

trends = np.array(trends, dtype=np.float32).transpose()
np.savetxt(end_time.split()[0]+'_'+str(space)+unit+'_'+str(n_samples)+'.csv', trends, delimiter=',')