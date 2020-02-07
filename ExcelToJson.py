import itertools
import threading
import time
import sys
import xlrd
import json 
import datetime


done = False
#here is the animation
def animate():
    for c in itertools.cycle(['|', '/', '-', '\\']):
        if done:
            break
        sys.stdout.write('\rloading ' + c)
        sys.stdout.flush()
        time.sleep(0.1)
    sys.stdout.write('\rDone!     ')

t = threading.Thread(target=animate)
t.start()


start = time.time()
#執行直接輸入路徑
file =sys.argv[1] 
data = xlrd.open_workbook(file)
table=data.sheets()[0]


def haveNoIndex(table):
    returnData=[]
    keyMap=table.row_values(0) 
    for i in range(table.nrows):#row
        tmpmp={}
        tmpInd=0
        for k in table.row_values(i): 
            tmpmp[keyMap[tmpInd]]=k
            tmpInd=tmpInd+1  
        returnData.append(tmpmp);
    return json.dumps(returnData,ensure_ascii=False,indent=2)

returnJson= haveNoIndex(table) 
fp = open(file+".json","w",encoding='utf-8')
fp.write(returnJson)
fp.close()

done = True

print("it took", time.time() - start, "seconds.")