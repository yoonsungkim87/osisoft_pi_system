from django.shortcuts import render
from django.utils import timezone
from .models import Post

# Create your views here.
def main(request):
    return render(request, 'app/main.html', {})

def chart(request):        
    import pythoncom
    import win32com.client as win32
    import numpy as np

    pythoncom.CoInitialize()
    server = win32.Dispatch('PISDK.PISDK.1').Servers('-')
    pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
    
    tag1 = '-'
    tag2 = '-'
    tag3 = '-'
    point1 = server.PIPoints(tag1).Data
    point2 = server.PIPoints(tag2).Data
    point3 = server.PIPoints(tag3).Data
    points = [point1,point2,point3]
    trends = []
    n_samples = 250
    for point in points:
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
        trends.append([tmpTime, tmpValue])
    pythoncom.CoUninitialize()
    
    import django
    from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
    from matplotlib import pyplot as plt
    from matplotlib.ticker import FormatStrFormatter
    import matplotlib.patches as mpatches
    import gc

    patch1 = mpatches.Patch(color='Blue', label=tag1)
    patch2 = mpatches.Patch(color='Green', label=tag2)
    patch3 = mpatches.Patch(color='Red', label=tag3)
    plt.legend(handles=[patch1,patch2,patch3], loc=2)
    with plt.style.context(u'seaborn-colorblind'):
        for trend in trends:
            plt.plot(np.array(trend)[0,:], np.array(trend)[1,:],'o-')
    plt.gca().yaxis.set_major_formatter(FormatStrFormatter('%d [%%]'))
    plt.gca().xaxis.set_major_formatter(FormatStrFormatter('%d [h]'))
    fig = plt.figure(1)
    fig.set_size_inches(16,10)

    canvas=FigureCanvas(fig)
    response=django.http.HttpResponse(content_type='image/png')
    canvas.print_png(response)
    
    fig.clf()
    plt.close()
    gc.collect()
    
    return response