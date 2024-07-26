import numpy as np
import pandas as pd
import re

Config=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Config")

Bulut_All=pd.read_excel(Config.iloc[4,1])
Banka_Conf=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Banka")
Bulut_All["Tarih"]=Bulut_All["Tarih"].dt.date

V3_Sablon=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="V3_Sablon")

Otoyol=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Otoyol")
Arac=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Arac")
HGS=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="HGS")

Banka_HGS = Bulut_All[Bulut_All['Fonksiyon Kodu 2'].isin(Otoyol['Fonksiyon'])].copy()


def extract_license_plate(text):
    match = re.search(r'\b[A-Z0-9]{2,3}[A-Z]{1,2}[0-9]{2,4}\b', text)
    return match.group(0) if match else None

Banka_HGS["Plaka"]=Banka_HGS["Açıklama"].apply(extract_license_plate)

Banka_HGS.rename(columns={"Fonksiyon Kodu 2": 'Fonksiyon'}, inplace=True)

DF_Merged=Banka_HGS[["İşlem Kodu","Banka","Tutar","Tarih","Plaka","Firma IBAN","Fonksiyon"]].merge(Otoyol, on="Fonksiyon", how='left')

DF_Merged["Açıklama"]=DF_Merged["Banka"]+"-"+DF_Merged["Plaka"]+"-"+DF_Merged["Açıklama"]


HGS_Data=DF_Merged.merge(Arac, on="Plaka", how="left")
HGS_Data.drop(['Banka', 'HGS_Etiket', "Fonksiyon", "Türü", "Plaka"], axis=1, inplace=True)

HGS_Data["KKEG"]=(-HGS_Data["Tutar"]*(1-HGS_Data["Gider"])).round(2)
HGS_Data["KKEG_Nazım_Y"]=-HGS_Data["KKEG"]
HGS_Data["KKEG_Nazım"]=HGS_Data["KKEG"]
HGS_Data["KDV_T"]=(((-HGS_Data["Tutar"])*HGS_Data["Gider"])/(1+HGS_Data["KDV"])*HGS_Data["KDV"]).round(2)
HGS_Data["Gider_T"]=(-HGS_Data["Tutar"])-HGS_Data["KDV_T"]-HGS_Data["KKEG"]
HGS_Data.rename(columns={"Tutar": 'Banka'}, inplace=True)

melted_df = pd.melt(HGS_Data, id_vars=["Tarih", "İşlem Kodu", "Firma IBAN", "Açıklama", "KDV"], value_vars=[ "Banka", "KKEG", "KKEG_Nazım_Y", "KKEG_Nazım","KDV_T", "Gider_T"], 
                    var_name="Type", value_name="Tutar").sort_values(by="İşlem Kodu")


melted_df.loc[(melted_df['Type'] == 'KDV_T') & (melted_df['KDV'] == 0.1), 'Type'] = 'KDV_10'
melted_df.loc[(melted_df['Type'] == 'KDV_T') & (melted_df['KDV'] == 0.2), 'Type'] = 'KDV_20'


V3_Data_1=melted_df[melted_df["Tutar"]!=0]

df=V3_Data_1.merge(HGS, on="Type", how="left")

V3_data=df.merge(Banka_Conf[["Firma IBAN","BANKA HESAP KODU", "Muhasebe Kodu"]], on="Firma IBAN", how="left")

V3_data.sort_values(by=["İşlem Kodu", "LineDocumentNumber"], inplace=True)

V3_data.reset_index(drop=True, inplace=True)

V3_Format=V3_data[["Tarih"]].copy()

V3_Format.rename(columns={"Tarih": 'DocumentDate'}, inplace=True)

V3_Format["DocumentNumber"]=V3_data["İşlem Kodu"]
V3_Format["Description"]=V3_data["Açıklama"]
V3_Format["BankCurrAccCode"] = np.where(V3_data["Type"] == "Banka", V3_data["BANKA HESAP KODU"], "")
V3_Format["BankOpTypeCode"]= 27
V3_Format["DueDate"]= ""
V3_Format["LineDescription"]= V3_data["Açıklama"]
V3_Format["LineDocumentNumber"]= V3_data["LineDocumentNumber"]
V3_Format["DocumentTypeCode"]= ""
V3_Format["DocumentTypeDescription"]= ""
V3_Format["PaymentMethod"]= "Banka"
V3_Format["GLAccCode"]= np.where(V3_data["Type"] == "Banka", V3_data["Muhasebe Kodu"], V3_data["Kod"])
V3_Format["CostCenterCode"]= "E1"
V3_Format["GLTypeCode"]= ""
V3_Format["ImportFileNumber"]= ""
V3_Format["ExportFileNumber"]= ""
V3_Format["DocCurrencyCode"]= "TRY"
V3_Format["ExchangeRate"]= 1
V3_Format["Debit"]= np.where(V3_data["Tutar"] > 0, V3_data["Tutar"], 0)
V3_Format["Credit"]= np.where(V3_data["Tutar"] < 0, -V3_data["Tutar"], 0)
V3_Format["ATAtt01"]= ""
V3_Format["ATAtt02"]= ""
V3_Format["ATAtt03"]= ""
V3_Format["ATAtt04"]= ""
V3_Format["ATAtt05"]= ""
V3_Format["FTAtt01"]= ""
V3_Format["FTAtt02"]= ""
V3_Format["FTAtt03"]= ""
V3_Format["FTAtt04"]= ""
V3_Format["FTAtt05"]= ""

V3_Format.to_excel(Config.iloc[7,1], index=False)