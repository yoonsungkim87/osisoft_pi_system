import pythoncom, pywintypes, os
import win32com.client as win32
import numpy as np

NUM_OF_SAMPLE = 12*24*30*2
SPACE = 5
UNIT = 'm'
END_TIME ='2021-06-26T00:00:00'
DELAY = ''  # -20s when end time is *
EXCPT = 'n' # r:reason, n:nan, b:blank
TAG_NAME_IN_RESULT = True

if EXCPT != 'r' and EXCPT != 'n' and EXCPT != 'b':
    raise('excpt type error!')

# Print iterations progress
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print('\r%s |%s| %s%% %s' % (prefix, bar, percent, suffix), end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()

server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)

point = []
iter_cnt = 1
err_cnt= 0
reason = set()

tags = [k for k in os.listdir() if 'tag' in k]
tag =np.array([])
for key in tags:
    tag = np.concatenate((tag,np.loadtxt(key, dtype=str, delimiter=',')))
tag = list(set(tag))
print(tag)

for x in tag:
    point.append(server.PIPoints(x).Data)
l = len(point)
trends = []

printProgressBar(0, l, prefix = 'Progress:', suffix = 'Complete', length = 50)

for i, p in enumerate(point):
    if p is not None:
        data2 = pisdk.IPIData2(p)
        #print('Extracting Data...')
        while True:
            try:
                results = data2.InterpolatedValues2(END_TIME+DELAY+'-'+str(int(NUM_OF_SAMPLE)*SPACE)+UNIT,END_TIME+DELAY,str(SPACE)+UNIT,asynchStatus=None)
                #print('**************************Successful!')
                break
            except pywintypes.com_error:
                #print('Error occured, retrying...')
                pass
        tmpValue = []
        tmpTime = []
        for v in results:
            try:
                if i == 0:
                    t = float(v.TimeStamp.LocalDate.timestamp())
                    tmpTime.append(t)
                s = str(v.Value)
                tmpValue.append(float(s))
            except ValueError:
                # if s == 'N RUN' or s == 'NRUN' or s == 'N OPEN' or s == 'NSTART' or s == 'OFF':
                if s in ['N RUN', 'NRUN', 'N OPEN', 'NSTART', 'OFF', 'N CLS', 'NO', 'N OPND', 'CLOSE', 'N OPN', 'NOPEND']:
                    tmpValue.append(0.0)
                # elif s == 'RUN' or s == 'OPEN' or s == 'START' or s == 'ON':
                elif s in ['RUN', 'OPEN', 'START', 'ON', 'OPENED', 'YES', 'OPEND',]:
                    tmpValue.append(1.0)
                else:
                    try:
                        if EXCPT == 'r':
                            tmpValue.append(s)
                        else:
                            tmpValue.append(np.nan)
                    #    tmpValue.append(tmpValue[-1])
                    #except IndexError:
                    #    tmpValue.append(0.0)
                    finally:
                        err_cnt += 1
                        reason.add(str(v.Value))
        if i == 0:
            #tmpTime.pop()
            trends.append(tmpTime)
        #tmpValue.pop()
        trends.append(tmpValue)
        printProgressBar(i + 1, l, prefix = 'Progress:', suffix = 'Complete', length = 50)
        
print('Total Error Counter: ', end='')
print(err_cnt)

print('Reason: ', end='')
print(*reason if reason else '', sep=', ')

trends = np.array(trends).transpose()
if EXCPT == 'b':
    trends = trends[~np.isnan(trends).any(axis=1)]
if TAG_NAME_IN_RESULT:
    tag.insert(0,'time')
    trends = np.concatenate((np.array(tag).reshape(1,-1),trends),axis=0 )

np.savetxt(END_TIME.split()[0].replace('*','crnt').replace(':','')+DELAY+'_'+str(SPACE)+UNIT+'_'+str(int(NUM_OF_SAMPLE))+'_'+EXCPT+'.csv', trends, delimiter=',', fmt='%s')
