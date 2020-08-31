import measurement as ms
import pandas as pd

def importFiles(filenames, mClass, nameparser):
    df = pd.DataFrame()
    for fname in filenames:
        data = mClass(fname).asDict()
        metadata = nameparser(fname)
        df = df.append({**metadata, **data}, ignore_index=True)
    return df

def sp(fname):
    name, _ = fname.split('.')
    name = name.split('/')[-1]
    params = name.split('_')
    s = params[0]
    c = params[1]
    a = params[2]
    if len(params) == 4:
        version = params[3]
        s += version
    return {'sample':s, 'capping':c, 'anneal':a}

# example script
import os
os.chdir('Anneal series ZnO Al + cap')

import glob
files = glob.glob('SE/*.SEsnap')
data = importFiles(files, ms.SEMeasurement, sp)

# pick the relevant stuff
shdata = data[['sample', 'capping', 'anneal','A2','B2','Thickness # 3']]
