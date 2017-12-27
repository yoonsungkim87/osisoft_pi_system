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
    'SE851HIS', #Turbine Speed
    'TE85147', #T1C
    'PT85425', #Fuel Gas Pressure
    'FM85404', #Starting Gas Valve
    'FM85403', #Main Gas Valve
    'FT85955', #WaterInj Flow
    'FM85310', #IGV
    
    'PT85168', #Inlet Scroll DP
    'PT85178S', #P2C
    'TE85315S', #T2C
    'JT86001S', #MW
    'TE8AVBP', #BP
    'TE8AVTX' #Exh
]

for x in tag:
    point.append(server.PIPoints(x).Data)
trends = []
n_samples = int(4*30*24*60)
space = 1
unit = 'm'
end_time = '2014-09-01 00:00'
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