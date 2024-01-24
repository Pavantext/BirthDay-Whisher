import pandas as pd 
import datetime
def sendEmail(to, sub, msg):
    print(f"Emial to '{to}' sent with sbject: '{sub}' and message '{msg}'")


if __name__ == "__main__":
    df = pd.read_excel("Data.xlsx")
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    # print(today)
    yearNow = datetime.datetime.now().strftime("%Y")
    # print(yearNow)
    writeInd = []
    for index, item in df.iterrows():
        # print(index, item["B.D"])
        bday = item['B.D'].strftime("%d-%m")
        # print(bday)
        if (today == bday) and yearNow not in str(item['Year']): # str() used cause 'Year' is in int frmt in Data file
            sendEmail(item['Email'], "Happy Birthday", item['Dialogue']) #Calling function with items of Data file
            writeInd.append(index)

    # print(writeInd)
    for i in writeInd:
        yr = df.loc[i, 'Year']
        # print(yr)
        df.loc[i,'Year'] = str(yr) + ',' + str(yearNow)
        # print(df.loc[i,'Year'])

    df.to_excel('Data.xlsx', index=False) 
        