# Prerequisite
- PI ProcessBook v3.2.3.0 ![PI](http://www.osisoft.com/images/osi-logo.png) 
- Django v1.10.1 ![django](https://avatars1.githubusercontent.com/u/27804?v=3&s=60)
- Anaconda v4.2.9 ![anaconda](https://www.continuum.io/sites/all/themes/continuum/assets/images/logos/logo-horizontal-large.svg)

# Main function is getting data from PI Server
Before you commit code, you need to specify server address and tag name. If you are not using Django framework, we should not do CoInitialize() / CoUninitialize().
```{.python}
import pythoncom
import win32com.client as win32
import numpy as np

pythoncom.CoInitialize()
server = win32.Dispatch('PISDK.PISDK.1').Servers(server_address)
pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
tag = tag_name
point = server.PIPoints(tag).Data
n_samples = 250
data2 = pisdk.IPIData2(point)
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
result_set = np.array([tmpTime, tmpValue], dtype=np.float64)
pythoncom.CoUninitialize()
```

# Result
![Result](https://github.com/yoonsungkim87/osisoft_pi_system/blob/master/trend.png)
