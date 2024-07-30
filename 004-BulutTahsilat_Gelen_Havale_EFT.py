import numpy as np
import pandas as pd

Config=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Config")
Banka_Conf=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Banka")
VKN_Conf=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="VKN", dtype={'Vergi No': str})

V3_Sablon=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="V3_Sablon")

# V3_Sablon[V3_Sablon["Belge No"]==3]["ColumName"].to_list()

Bulut_All=pd.read_excel(Config.iloc[4,1], dtype={'Vergi No': str})
Bulut_All["Tarih"]=Bulut_All["Tarih"].dt.date

Bulut_Filtered = Bulut_All[
    (Bulut_All['Vergi No'].isin(VKN_Conf['Vergi No'])) |
    ((Bulut_All['Banka'] == 'TEB') & Bulut_All['Açıklama'].str.contains("nolu fat") & 
     Bulut_All['Açıklama'].str.contains("L J MAĞAZACILIK SAN VE TİİV AŞ")&
     (Bulut_All['İşlem Tipi'] == 'BANKA HAREKETİ') 
   )].copy().reset_index(drop=True)


Bulut_Filtered.loc[((Bulut_Filtered['Banka'] == 'TEB') & Bulut_Filtered['Açıklama'].str.contains("nolu fat") & 
     Bulut_Filtered['Açıklama'].str.contains("L J MAĞAZACILIK SAN VE TİİV AŞ")&
     (Bulut_Filtered['İşlem Tipi'] == 'BANKA HAREKETİ') 
   ), 'Vergi No'] = '6080051647'

Bank_Code=Bulut_Filtered.merge(Banka_Conf[["Firma IBAN","BANKA HESAP KODU", "Muhasebe Kodu"]], on="Firma IBAN", how="left")

V3_data=Bank_Code.merge(VKN_Conf, on="Vergi No", how="left").sort_values(by="İşlem Kodu")

V3_Format=V3_data[["Tarih"]].copy()
V3_Format.rename(columns={"Tarih": 'DocumentDate'}, inplace=True)

V3_Format["DocumentNumber"]=V3_data["İşlem Kodu"]
V3_Format["Description"]=V3_data["Banka"]+"-"+V3_data["Açıklama_y"]+"-"+np.where(V3_data["Tutar"] > 0, "TAHSİLATI", "ÖDEMESİ")
V3_Format["BankCurrAccCode"] = V3_data["BANKA HESAP KODU"]
V3_Format['BankTransTypeCode'] = np.where(V3_data["Tutar"] > 0, 4, 5)
V3_Format['CurrAccTypeCode']= V3_data['CurrAccTypeCode']
V3_Format['CurrAccCode']= V3_data['CurrAccCode']
V3_Format['SubCurrAccCode']= ""
V3_Format['CurrAccCurrencyCode']= "TRY"
V3_Format['CurrAccExchangeRate']= 1
V3_Format['DocCurrencyCode']= "TRY"
V3_Format['ExchangeRate'] = 1
V3_Format['Doc_Amount']= V3_data['Tutar'].abs()
V3_Format['Doc_TransferCharges'] = ""
V3_Format['GLTypeCode'] = ""
V3_Format['ImportFileNumber'] = ""
V3_Format['ExportFileNumber'] = ""
V3_Format['LineDescription'] = V3_data["Banka"]+"-"+V3_data["Açıklama_y"]+"-"+np.where(V3_data["Tutar"] > 0, "TAHSİLATI", "ÖDEMESİ")
V3_Format['LineDocumentNumber'] = 1
V3_Format['ATAtt01'] = ""
V3_Format['ATAtt02'] = ""
V3_Format['ATAtt03'] = ""
V3_Format['ATAtt04'] = ""
V3_Format['ATAtt05'] = ""
V3_Format['FTAtt01'] = ""
V3_Format['FTAtt02'] = ""
V3_Format['FTAtt03'] = ""
V3_Format['FTAtt04'] = ""
V3_Format['FTAtt05'] = ""

V3_Format.to_excel(Config.iloc[13,1], index=False)