import xlwt
import xlrd
import os

def GetPngDatas():
    PngDatalist = []
    pngsPath = "ChatEmojis"
    filelist = os.listdir(pngsPath)
    for item in filelist:
        pngData = {}
        pngData["Id"] = os.path.splitext(item)[0]
        resourceStr ="ResourceObject=Texture2D\'\"/Game/UI/Images/NoPacker/Emojis/" + pngData["Id"] + "." +pngData["Id"] + "\"\',"
        pngData["Brush"] = "(ImageSize=(X=72,Y=72),Margin=(Left=0.000000,Top=0.000000,Right=0.000000,Bottom=0.000000),TintColor=(SpecifiedColor=(R=1.000000,G=1.000000,B=1.000000,A=1.000000),ColorUseRule=UseColor_Specified)," + resourceStr + "ResourceName=\"\",UVRegion=(Min=(X=0.000000,Y=0.000000),Max=(X=0.000000,Y=0.000000),bIsValid=0),DrawAs=Image,Tiling=NoTile,Mirroring=NoMirror,ImageType=NoImage,bIsDynamicallyLoaded=False"
        PngDatalist.append(pngData)
    return PngDatalist
def saveDataToExcel():
    mylist = GetPngDatas()
    excelName = "EmojiData.xls"
    file = xlwt.Workbook(encoding = "utf-8")
    sheet = file.add_sheet(u'data')
    sheet.write(0,0,"")
    sheet.write(0,1,"Brush")
    headers = list(mylist[0].keys())
    for i in range(1,len(mylist)+1):
        for j in range(len(mylist[i-1])):
            sheet.write(i,j,mylist[i-1][headers[j]])
    file.save(excelName)
    
saveDataToExcel()
