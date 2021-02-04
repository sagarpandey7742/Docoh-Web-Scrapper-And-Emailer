from util import *
import os
from csv import writer

driver = webdriver.Firefox(executable_path=driverPath, options=options)

rowFound = False
tickerr = []
fundr = []
formr = []
purposer = []
companyNamer = []
timeStampr = []
linkr = []
emailBodyr = []
longTextr = []

urls = ["https://docoh.com/search?q=8-K%20AND%20%22Material%20Definitive%20Agreement%22%20NOT%20(" \
        "%20%22Partner%22%20OR%20%22LP%22%20OR%20%22Compensation%22%20OR%20%22exercise%22%20)&page=",
        "https://docoh.com/search?q=8-K%20AND%20%E2%80%9CStock%20Split%E2%80%9D&page=",
        "https://docoh.com/search?q=SC%2013D&page=",
        "https://docoh.com/search?q=SC%2013G&page=",
        "https://docoh.com/search?q=13F-HR&page="]

dataset = set()
datalist = []
if os.path.isfile(outputPath + "output.csv"):
    dataset, datalist = generate_list(set(), [])

print(len(dataset), len(datalist))
for url in enumerate(urls):
    rowFound = False
    for page in range(1, 10):
        driver.get(url[1] + str(page))
        for i in range(1, 21):
            row = driver.find_element_by_xpath(
                '//*[@id="main-col"]/div[2]/div/div[2]/div/div/div[3]/div/a[' + str(i) + "]")
            ticker, fund, form, purpose, companyName, timeStamp, link, emailBody, longText = getElements(row, i,
                                                                                                         url[0] + 1)
            if os.path.isfile(outputPath + "output.csv"):
                rowFound = checkInCsv(dataset, datalist, ticker, fund, form, purpose, companyName, timeStamp, link,
                                      emailBody, longText, rowFound)
            if rowFound:
                break
            tickerr.append(ticker)
            fundr.append(fund)
            purposer.append(purpose)
            companyNamer.append(companyName)
            timeStampr.append(timeStamp)
            linkr.append(link)
            formr.append(form)
            emailBodyr.append(emailBody)
            longTextr.append(longText)
        if rowFound == True:
            break

if not os.path.isfile(outputPath + "output.csv"):
    name_of_file = "output"
    completeName = os.path.join(outputPath + name_of_file + ".csv")

    data = pd.DataFrame({
        "Form": formr,
        "Ticker": tickerr,
        "CompanyName": companyNamer,
        "Fund": fundr,
        "Purpose": purposer,
        "EmailBody": emailBodyr,
        "LongText": longTextr,
        "Link": linkr,
        "TimeStamp": timeStampr
    })

    data.to_csv(completeName, index=False, header=False)
    print("Main file written")

else:
    print(len(datalist))
    tickerr = []
    fundr = []
    formr = []
    purposer = []
    companyNamer = []
    timeStampr = []
    linkr = []
    emailBodyr = []
    longTextr = []
    for i in range(0, len(datalist)):
        formr.append(datalist[i][0])
        tickerr.append(datalist[i][1])
        companyNamer.append(datalist[i][2])
        fundr.append(datalist[i][3])
        purposer.append(datalist[i][4])
        emailBodyr.append(datalist[i][5])
        longTextr.append(datalist[i][6])
        linkr.append(datalist[i][7])
        timeStampr.append(datalist[i][8])

    name_of_file = "output"
    completeName = os.path.join(outputPath + name_of_file + ".csv")
    data = pd.DataFrame({
        "Form": formr,
        "Ticker": tickerr,
        "CompanyName": companyNamer,
        "Fund": fundr,
        "Purpose": purposer,
        "EmailBody": emailBodyr,
        "LongText": longTextr,
        "Link": linkr,
        "TimeStamp": timeStampr
    })

    data.to_csv(completeName, index=False, header=False)

driver.close()
