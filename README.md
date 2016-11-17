# This code is based on OSI PI ProcessBook & django
![PI](http://www.osisoft.com/images/osi-logo.png) ![django](https://avatars1.githubusercontent.com/u/27804?v=3&s=60)

# Main function is getting data from PI Server
[[TIP]]
Before you commit code, you need to specify server address and tag name. If you are not using Django framework, we should not do CoInitialize() / CoUninitialize().
[[/TIP]]
```{.py}
pythoncom.CoInitialize()
server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
    
tag, point = [None] * 3, [None] * 3
tag[0] = '70MKG31CQ001A_XQ01'
tag[1] = '80MKG31CQ001A_XQ01'
tag[2] = '90MKG31CQ001A_XQ01'
point[0] = server.PIPoints(tag[0]).Data
point[1] = server.PIPoints(tag[1]).Data
point[2] = server.PIPoints(tag[2]).Data
trends = []
n_samples = 250
for p in point:
    data2 = pisdk.IPIData2(p)
    results = data2.InterpolatedValues2('*-'+str(n_samples)+'h','*','1h',asynchStatus=None)
    tmpValue =[]
    tmpTime = []
    i = 1 - n_samples
    for v in results:
        try:
            s = str(v.Value)
            tmpValue.append(float(s))
            tmpTime.append(i)
        except ValueError:
            pass
        i = i + 1
    tmpValue.pop()
    tmpTime.pop()
    trends.append([tmpTime, tmpValue])
pythoncom.CoUninitialize()
```
