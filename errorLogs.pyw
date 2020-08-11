import traceback
import os
import datetime

def errorLog(func):

    def wrapper(*args, **kwargs):

        try:
            return func(*args, **kwargs)
        except Exception as exc:
            now = datetime.datetime.now()
            if not os.path.exists(r'.\errorLogs'):
                os.makedirs(r'.\errorLogs')
            logFile = open(r'.\errorLogs\tool_logfile_%s.txt' %'_'.join([str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second)]),
                           'w')
            logFile.write(traceback.format_exc())
            logFile.close()
            print('Exception written to file')
            raise Exception(exc)

    return wrapper
