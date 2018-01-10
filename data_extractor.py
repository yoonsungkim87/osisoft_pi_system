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
    '80MBY10CS901_ZQ11', #Turbine Speed
    '80MBA11CT901_ZQ01', #T1C
    '80MBA11CP101_XQ01', #P1C
    '80MBP13FP901_ZQ02', #Fuel Gas Pressure
    '80MBP21FF901_ZQ01', # A Stage Gas Flow
    '80MBP22FF901_ZQ01', # B Stage Gas Flow
    '80MBP23FF901_ZQ01', # C Stage Gas Flow
    '80MBP32FF901_ZQ01', # D Stage Gas Flow
    '80MBP31FF901_ZQ01', # P Stage Gas Flow
    '70MBP13AA001_ZB21', #Main ESV
    '80MBA11DG011_XQ01', #IGV
    '80MBA10FG100_ZV01', #HCO On
    
    '80CJA00FF001_ZQ02', #Air Flow
    '80MBA12CP901_ZQ01', #P2C
    '80MBA12CT901_ZQ01', #T2C
    '80MBY10CE901_XQ01', #MW
    '80MBA28CT900_ZQ01', #BP
    '80MBR10CT900_ZQ01' #Exh
]

for x in tag:
    point.append(server.PIPoints(x).Data)
trends = []
n_samples = int(4*30*24*60)
space = 1
unit = 'm'
end_time = '2015-09-01 00:00'
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