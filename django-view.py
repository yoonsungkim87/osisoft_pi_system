def chart(request):        
    import pythoncom
    import win32com.client as win32
    import numpy as np

    #put your server ip address
    srv_ip = ''
    pythoncom.CoInitialize()
    server = win32.Dispatch('PISDK.PISDK.1').Servers('POSCOPOWER')
    pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
    
    tag, point = [None] * 3, [None] * 3
    #tag point should be indicated. this code is based upon 3 tags.
    tag[0] = ''
    tag[1] = ''
    tag[2] = ''
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
    
    import django
    from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
    from matplotlib import pyplot as plt
    from matplotlib.ticker import FormatStrFormatter
    import matplotlib.patches as mpatches
    import gc

    i = 0
    with plt.style.context(u'seaborn-colorblind'):
        for trend in trends:
            plt.plot(np.array(trend[0]), np.array(trend[1]),'o-', label = tag[i])
            i += 1
    plt.legend(loc=2)
    
    plt.gca().yaxis.set_major_formatter(FormatStrFormatter('%d [%%]'))
    plt.gca().xaxis.set_major_formatter(FormatStrFormatter('%d [h]'))
    fig = plt.figure(1)
    fig.set_size_inches(16,10)

    canvas=FigureCanvas(fig)
    response=django.http.HttpResponse(content_type='image/png')
    canvas.print_png(response)
    
    plt.close('all')
    gc.collect()
    
    return response