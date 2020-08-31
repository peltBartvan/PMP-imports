import zipfile
import os
import warnings
import pandas as pd

class Measurement():
    # abstract class
    # no functionality here
    # simply serves to raise errors when user doesn't define proper subclasses
    # may upgrade to metaclass when I feel like it, seems interesting!
    def __init__(self, path):
        # delegate automatically to proper subclass
        raise NotImplementedError()

    def __getitem__(self, key):
        # use meas['feature'] to get the proper feature
        raise NotImplementedError()

    def keys(self):
        raise NotImplementedError()

    def asDict(self, *a, **kw):
        if len(a) == 0:
            keys = self.keys()
        else:
            keys = list(a)
        return {key: self[key] for key in keys}

class ExcelMeasurement(Measurement):
    def __init__(self, path):
        if path.endswith('.xlsm'):
            with warnings.catch_warnings():
                warnings.simplefilter('ignore')
                self.df = pd.read_excel(path,
                        sheet_name = None,
                        header = None,
                        engine = 'openpyxl')
        elif path.endswith('.xlsx'):
            self.df = pd.read_excel(path,
                    sheet_name = None,
                    header = None)
        else:
            raise ValueError('File must be a .xlsm or .xlsx file')

    def __getitem__(self, key):
        try:
            sheet, colname, row = self.cellDict[key]
            return self.df[sheet][colname][row]
        except KeyError:
            print(key + ' is not a valid parameter')
            return None

    def keys(self):
        return self.cellDict.keys()

class SintonMeasurement(ExcelMeasurement):
    cellDict = {
            "Sample Name"    : ["User", 0, 5],
            "Wafer Thickness": ["User", 1, 5],
            "Resistivity"    : ["User", 2, 5],
            "Sample Type"    : ["User", 3, 5],

            "Optical Constant" : ["User", 4, 5],
            "Specified MCD"    : ["User", 5, 5],
            "Bias Light"       : ["User", 6, 5],
            "Analysis Mode"    : ["User", 7, 5],
            "Lifetime at Spec. MCD" : ["User", 0, 8],
            "Sheet Resistance"      : ["User", 1, 8],
            "Measured Resistivity"  : ["User", 2, 8],
            "J0"                    : ["User", 3, 8],
            "Fit Intercept"         : ["User", 4, 8],
            "Min MCD"               : ["User", 5, 8],
            "Max MCD"               : ["User", 6, 8],
            "Bias point CD"         : ["User", 7, 8],
            "Trap Density"          : ["User", 8, 8],
            "Doping"                : ["User", 9, 8],
            "1 sun Implied Voc"     : ["User", 10, 8],
            "Implied FF"            : ["User", 11, 8],

            "lifetime" : ["User", 0, 8],
            "iVoc"     : ["User", 10, 8],
            "iFF"      : ["User", 11, 8],
            }

    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        self.sanityCheck()

    def sanityCheck(self):
        setRes = self['Resistivity']
        measRes = self['Measured Resistivity']
        relerror = abs((setRes-measRes)/measRes)
        threshold = 0.01
        if relerror > threshold:
            print("Warning: resistivities do not match in sample '" + self['Sample Name']+"'")

class SEMeasurement(Measurement):
    def __init__(self, path):
        self.path = path
        logstring = self.getLogString(path)
        self.valueDict = self.extractData(logstring)

    def getLogString(self, path):
        zf = zipfile.ZipFile(path)
        fitlog = zf.extract('_FitLog')
        with open(fitlog, encoding='latin-1') as f:
            logstring = f.read()
        os.remove(fitlog)
        return logstring

    def readBetween(self, s, start, end):
        return s[s.find(start)+len(start):s.rfind(end)]

    def extractData(self, logstring):
        startstr = 'start_Fit Parms'
        endstr   = 'end_Fit Parms'
        s = self.readBetween(logstring, startstr, endstr)
        ssplit = s.replace('\t','').replace("'",'').split('\n')[2:-2]
        valdict = {key: value for (key, value) in map(lambda x: (x[0].strip(), float(x[-1])),
            map(lambda x: x.split('='), ssplit))}
        return valdict

    def __getitem__(self, key):
        return self.valueDict[key]

    def keys(self):
        return self.valueDict.keys()

class HallMeasurement(ExcelMeasurement):
    cellDict = {
            "Hall mobility"                 : ['Summary', 4, 21],
            "Carrier type"                  : ['Summary', 4, 22],
            "Carrier concentration"         : ['Summary', 4, 23],
            "Sheet carrier concentration"   : ['Summary', 4, 24],
            "Hall coefficient"              : ['Summary', 4, 25],
            "Sheet Hall coefficient"        : ['Summary', 4, 26],
            "Resistivity"                   : ['Summary', 4, 27],
            "Sheet resistivity"             : ['Summary', 4, 28],
            "Hall voltage"                  : ['Summary', 4, 29],
            "Thickness"                     : ['Summary', 2,  9],
            }

path = '~/Documents/Msc/metingen'
path = os.path.expanduser(path)
os.chdir(path)
