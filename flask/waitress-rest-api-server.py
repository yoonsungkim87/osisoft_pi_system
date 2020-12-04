from flask import Flask
from flask_cors import CORS
from waitress import serve
from flask_restful import reqparse, abort, Api, Resource

import pythoncom, pywintypes, requests, json, datetime, os
import win32com.client as win32
#import numpy as np
#import tensorflow as tf

app = Flask(__name__)
CORS(app)
api = Api(app)

#tf.compat.v1.logging.set_verbosity(tf.compat.v1.logging.ERROR)

class TagsForKeyword(Resource):
    def get(self, keyword):
        # preventing overhead on server
        #if len(keyword) < 4:
        #    abort(404, error="Keyword '{}' is too short to search.".format(keyword))
            
        pythoncom.CoInitialize()
        server = win32.Dispatch('PISDK.PISDK').Servers('POSCOPOWER')
        result = {}
        #print(server.PIPoints('PTB5504').PointAttributes('descriptor').Value)
        ptList = server.GetPoints("tag = '"+keyword+"'",None)
        
        # abort if there are no results.
        if ptList.Count == 0:
            abort(404, error="No results for Keyword '{}'.".format(keyword))
            
        for i in ptList:
            result[str(i.Name)] = {
                'descriptor':str(i.PointAttributes('descriptor').Value),
                'value':str(i.Data.Snapshot),
                'engunits':str(i.PointAttributes('engunits').Value),
                'timestamp':str(i.Data.Snapshot.TimeStamp.LocalDate)
            }
        pythoncom.CoUninitialize()
        
        return result
    
class GroupLiveTags(Resource):
    def get(self, tags):
        
        tags= tags.split(",")
        pythoncom.CoInitialize()
        server = win32.Dispatch('PISDK.PISDK').Servers('POSCOPOWER')
        
        result = {}
        for tag in tags:
            try:
                result[tag] = [str(server.PIPoints(tag).Data.Snapshot.TimeStamp.LocalDate), str(server.PIPoints(tag).Data.Snapshot)]
            except pywintypes.com_error as e:
                abort(404, error="{} with {}".format(repr(e), tag))
        pythoncom.CoUninitialize()
        
        return result

class GroupRecordedTags(Resource):
    def get(self, period, tags, delay):
        
        tags= tags.split(",")
        pythoncom.CoInitialize()
        server = win32.Dispatch('PISDK.PISDK').Servers('POSCOPOWER')
        pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
        result = {}
        for tag in tags:
            try:
                data2 = pisdk.IPIData2(server.PIPoints(tag).Data)
                pitmp = data2.RecordedValues('*-'+period+delay,'*'+delay,3,"",0,None)
            except pywintypes.com_error as e:
                abort(404, error="{}".format(repr(e)))
            tmp = []
            for val in pitmp:
                v = str(val.Value)
                t = str(val.TimeStamp.LocalDate)
                tmp.append([t, v])
            #tmp.pop()
            result[tag] = tmp
        pythoncom.CoUninitialize()
        
        return result
    
class GroupIPRecordedTags(Resource):
    def get(self, period, freq, tags, delay):
        
        tags= tags.split(",")
        pythoncom.CoInitialize()
        server = win32.Dispatch('PISDK.PISDK').Servers('POSCOPOWER')
        pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
        result = {}
        for tag in tags:
            try:
                data2 = pisdk.IPIData2(server.PIPoints(tag).Data)
                pitmp = data2.InterpolatedValues2('*-'+period+delay,'*'+delay,freq,asynchStatus=None)
            except pywintypes.com_error as e:
                abort(404, error="{}".format(repr(e)))
            tmp = []
            for val in pitmp:
                v = str(val.Value)
                t = str(val.TimeStamp.LocalDate)
                tmp.append([t, v])
            #tmp.pop()
            result[tag] = tmp
        pythoncom.CoUninitialize()
        
        return result

'''
class PredictPondLevel(Resource):
    def get(self, variables, curr_pond_lvl, pred_sea_lvl):

        variables = variables.split(",")
        pred_sea_lvl = pred_sea_lvl.split(",")

        mm = np.loadtxt('maxmin.csv', delimiter=",", dtype=np.float64)
        model = tf.keras.models.load_model('myModel.h5')
        prd_lgth = 24
        lb = 0
        x_attr = (1,16,2,3,4,5,18,19,14,15,)
        y_attr = (21,)
        comp_max = mm[0,:]
        comp_min = mm[1,:]

        c1 = np.array([curr_pond_lvl] + [np.nan]*prd_lgth, dtype=np.float64).reshape(-1,1)
        c2 = np.array(pred_sea_lvl, dtype=np.float64).reshape(-1,1)
        r1 = np.array(variables[0::4]*6, dtype=np.float64).reshape(6,-1)
        r2 = np.array(variables[1::4]*6, dtype=np.float64).reshape(6,-1)
        r3 = np.array(variables[2::4]*6, dtype=np.float64).reshape(6,-1)
        r4 = np.array(variables[3::4]*7, dtype=np.float64).reshape(7,-1)
        c3 = np.concatenate((r1,r2,r3,r4), axis=0)
        var_temp = np.concatenate((c1,c2,c3), axis=1)   
        var_temp2 = ((var_temp - comp_min[None,:len(x_attr)]) / (comp_max - comp_min)[None,:len(x_attr)])

        true_temp = var_temp2.reshape(1,lb + prd_lgth + 1,-1)
        pred_temp = var_temp2[:1,:].reshape(1,1,-1)

        i = 0
        while pred_temp.shape[1] < lb + prd_lgth + 1:
            pred_temp2 = model.predict(pred_temp[:,-lb-1:,:])
            pred_temp3 = pred_temp2 * (comp_max - comp_min)[None,-len(y_attr):] + comp_min[None,-len(y_attr):]
            pred_temp4 = pred_temp[:,-1,:len(y_attr)] * (comp_max - comp_min)[None,:len(y_attr)] + comp_min[None,:len(y_attr)]
            pred_temp5 = ((pred_temp3 + pred_temp4 - comp_min[None,:len(y_attr)]) / (comp_max - comp_min)[None,:len(y_attr)]).reshape(-1,1,len(y_attr))
            true_temp2 = true_temp[:,lb+1+i:lb+2+i,len(y_attr):]
            pred_temp6 = np.concatenate((pred_temp5, true_temp2), axis = 2)
            pred_temp = np.concatenate((pred_temp, pred_temp6), axis = 1)
            pred_temp2 = None
            pred_temp3 = None
            pred_temp4 = None
            pred_temp5 = None
            pred_temp6 = None
            true_temp2 = None
            i = i + 1

        pred_ = pred_temp * (comp_max - comp_min)[None,:len(x_attr)] + comp_min[None,:len(x_attr)]
            
        return pred_[0,:,0].tolist()
'''
    
api.add_resource(TagsForKeyword, '/tags-for-keyword/<string:keyword>')
api.add_resource(GroupLiveTags, '/group-live-tags/<string:tags>')
api.add_resource(GroupRecordedTags, '/group-recorded-tags/<string:period>/<string:tags>/<string:delay>')
api.add_resource(GroupIPRecordedTags, '/group-ip-recorded-tags/<string:freq>/<string:period>/<string:tags>/<string:delay>')
#api.add_resource(PredictPondLevel, '/predict-pond-level/<string:variables>/<string:curr_pond_lvl>/<string:pred_sea_lvl>')

## Serve using waitress
serve(app, host='me', port=8080)