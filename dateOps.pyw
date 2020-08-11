import datetime

def getNetworkDay(differenceInDays,
                  delimiter='',
                  startdate=datetime.datetime.now(),
                  reverse=False):

    daysElapsed = 0
    try:
        step = int(differenceInDays/abs(differenceInDays))
    except:
        differenceInDays = -1
        step = int(differenceInDays/abs(differenceInDays))

    date = startdate
    while not daysElapsed == abs(differenceInDays):
        date = date + datetime.timedelta(days=step)
        daysElapsed += 1
        while date.weekday() in [5, 6]:
            date = date + datetime.timedelta(days=step)
            

    year = date.isoformat()[:4]
    month = date.isoformat()[5:7]
    day = date.isoformat()[8:10]

    if reverse:
        return delimiter.join([year, month, day])
    else:
        return delimiter.join([day, month, year])
    

def getExpiryRange():

    endDatetime = datetime.datetime.now() - datetime.timedelta(days=1)

    endYear = endDatetime.isoformat()[:4]
    endMonth = endDatetime.isoformat()[5:7]
    endDay = endDatetime.isoformat()[8:10]

    startDate = '01012015'
    endDate = ''.join([endDay, endMonth, endYear])

    return(startDate, endDate)

def isValidDate(string):
    try:
        day = string[:2]
        month = string[2:4]
        year = string[4:8]
    except:
        return False

    try :
        datetime.datetime(day = int(day), month = int(month), year = int(year))
        return True
    except ValueError :
        return False
    

