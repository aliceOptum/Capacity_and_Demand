import pandas as pd

def calculatePoint(size):
    if size == ("F XXS"):
        return(5)
    elif size == ("F XS"):
        return(13)
    elif size == ("F S"):
        return(20)
    elif size == ("F M"):
        return(40)
    elif size == ("F L"):
        return(80)
    elif size == ("F XL"):
        return(160)
    else:
        return(50)
def filter(name):
    try:
        x = name.split(" ")
        output = x[2]
        return(output[:-1])
    except:
        return("error")


xls = pd.ExcelFile('testingSort.xlsx')
export = pd.read_excel(xls, 'export')
capSheet = pd.read_excel(xls, 'Capabilities')
feat = pd.read_excel(xls, 'Features')

length = len(export.index)

export['points'] = [0 for i in range(length)]


for i,j in feat.iterrows():
    point = calculatePoint(j['Preliminary Estimate'])
    cap = filter(j['Parent'])
    capRow = capSheet.loc[capSheet['Formatted ID']== cap]
    try:
        st = (filter(capRow['Parent'].values[0]))
    except:
        st = "Ignore"
    """
    print (j['Formatted ID'])
    print(st)"""

    if st != "Ignore":
        stRow = export.loc[export['Formatted ID']== st]
        try:
            index = (stRow.index).values[0]
            old = export['points'][index]
            print("this is old")
            print(old)
            new = old + point
            print(new)
            
        
            export.loc[index, 'points'] =  new
            print("THIS IS NEW")
            print(export['points'][index])
    
   
        except:
            print("EHERERER")
            old = 0
       



export.to_csv('new.csv')
        


  





