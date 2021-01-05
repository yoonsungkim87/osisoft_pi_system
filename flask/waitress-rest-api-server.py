from flask import Flask
from flask_cors import CORS
from waitress import serve
from flask_restful import reqparse, abort, Api, Resource

import pythoncom, pywintypes, requests, json, datetime, os
from dateutil.parser import parse
import win32com.client as win32

app = Flask(__name__)
CORS(app)
api = Api(app)

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
            except Exception as e:
                abort(404, error="{} with {}".format(repr(e), tag))
        pythoncom.CoUninitialize()
        
        return result

class GroupRecordedTags(Resource):
    def get(self, tags, end, period, delay):
        
        tags = tags.split(",")
        try:
            end = parse(end)
            period = parse(period)-parse('0s')
            delay = parse(delay)-parse('0s')
        except Exception as e:
            abort(404, error="{}".format(repr(e)))
        pythoncom.CoInitialize()
        server = win32.Dispatch('PISDK.PISDK').Servers('POSCOPOWER')
        pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
        result = {}
        for tag in tags:
            try:
                data2 = pisdk.IPIData2(server.PIPoints(tag).Data)
                pitmp = data2.RecordedValues(str(end-period-delay),str(end-delay),3,"",0,None)
            except Exception as e:
                abort(404, error="{}".format(repr(e)))
            tmp = []
            for val in pitmp:
                v = str(val.Value)
                t = str(val.TimeStamp.LocalDate)
                tmp.append([t, v])
            result[tag] = tmp
        pythoncom.CoUninitialize()
        
        return result
    
class GroupIPRecordedTags(Resource):
    def get(self, tags, freq, end, period, delay):
        
        tags= tags.split(",")
        try:
            end = parse(end)
            period = parse(period)-parse('0s')
            delay = parse(delay)-parse('0s')
        except Exception as e:
            abort(404, error="{}".format(repr(e)))
        pythoncom.CoInitialize()
        server = win32.Dispatch('PISDK.PISDK').Servers('POSCOPOWER')
        pisdk = win32.gencache.EnsureModule('{0EE075CE-8C31-11D1-BD73-0060B0290178}',0, 1, 1,bForDemand = False)
        result = {}
        for tag in tags:
            try:
                data2 = pisdk.IPIData2(server.PIPoints(tag).Data)
                pitmp = data2.InterpolatedValues2(str(end-period-delay),str(end-delay),freq,asynchStatus=None)
            except Exception as e:
                abort(404, error="{}".format(repr(e)))
            tmp = []
            for val in pitmp:
                v = str(val.Value)
                t = str(val.TimeStamp.LocalDate)
                tmp.append([t, v])
            result[tag] = tmp
        pythoncom.CoUninitialize()
        
        return result
    
api.add_resource(TagsForKeyword, '/tags-for-keyword/<string:keyword>')
api.add_resource(GroupLiveTags, '/group-live-tags/<string:tags>')
api.add_resource(GroupRecordedTags, '/group-recorded-tags/<string:tags>/<string:end>/<string:period>/<string:delay>')
api.add_resource(GroupIPRecordedTags, '/group-ip-recorded-tags/<string:tags>/<string:freq>/<string:end>/<string:period>/<string:delay>')

## Serve using waitress
serve(app, host='me', port=8080)