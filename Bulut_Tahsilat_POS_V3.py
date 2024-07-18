import numpy as np
import pandas as pd

Config=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Config")

Bulut_All=pd.read_excel(Config.iloc[4,1])
Banka_Conf=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Banka")
Bulut_All["Tarih"]=Bulut_All["Tarih"].dt.date

V3_Sablon=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="V3_Sablon")

V3_Sablon_2=V3_Sablon[V3_Sablon["Belge No"]==2]

Bulut_POS = Bulut_All[(Bulut_All["İşlem Tipi"] == "POS") & 
~((Bulut_All["Banka"] == "HALKBANK") & Bulut_All["Açıklama"].str.contains("POS Aidat"))&
~(Bulut_All["Banka"] == "GARANTİ BBVA")&
~(Bulut_All["Banka"] == "İŞ BANKASI")].copy()

def determine_kod(row):
    if row['Banka'] == 'ZİRAAT BANKASI' and row['Hareket Tipi'] == 'ALACAK':
        return 'ZB-P'
    elif row['Banka'] == 'ZİRAAT BANKASI' and row['Hareket Tipi'] == 'BORÇ':
        return 'ZB-PI'
    elif row['Banka'] == 'HALKBANK' and row['Fonksiyon Kodu 1'] == 'CCP':
        return 'HB-P'
    elif row['Banka'] == 'HALKBANK' and row['Fonksiyon Kodu 1'] == 'MSC':
        return 'HB-PI'
    elif row['Banka'] == 'HALKBANK' and row['Fonksiyon Kodu 1'] == 'COC' and 'Aidat' in row['Açıklama']:
        return 'HB-M'
    elif row['Banka'] == 'HALKBANK' and row['Fonksiyon Kodu 1'] == 'COC' and 'Aidat' not in row['Açıklama']:
        return 'HB-K'
    elif row['Banka'] == 'VAKIFBANK' and row['Fonksiyon Kodu 1'] == 'CCP':
        return 'VB-P'
    elif row['Banka'] == 'VAKIFBANK' and row['Fonksiyon Kodu 2'] == 'UPSUYEBLOKE8':
        return 'VB-K'
    elif row['Banka'] == 'VAKIFBANK' and row['Fonksiyon Kodu 2'] == 'UPSUYEBLOKE7':
        return 'VB-PI'
    elif row['Banka'] == 'VAKIFBANK' and row['Fonksiyon Kodu 2'] == 'UPSUYEBLOKE2':
        return 'VB-KI'
    elif row['Banka'] == 'AKBANK' and row['Fonksiyon Kodu 1'] == 'CCP':
        return 'AK-P'
    elif row['Banka'] == 'AKBANK' and row['Fonksiyon Kodu 1'] == 'COC':
        return 'AK-K'
    elif row['Banka'] == 'QNB FİNANSBANK' and row['Fonksiyon Kodu 1'] == 'CCP':
        return 'QB-P'
    elif row['Banka'] == 'QNB FİNANSBANK' and row['Fonksiyon Kodu 1'] == 'COM' and '(Satış)' in row['Açıklama']:
        return 'QB-K'
    elif row['Banka'] == 'QNB FİNANSBANK' and row['Fonksiyon Kodu 1'] == 'COM' and '(İade' in row['Açıklama']:
        return 'QB-PI'
    elif row['Banka'] == 'QNB FİNANSBANK' and row['Fonksiyon Kodu 1'] == 'MSC' and row['Tutar'] > 0:
        return 'QB-P'
    elif row['Banka'] == 'QNB FİNANSBANK' and row['Fonksiyon Kodu 1'] == 'MSC' and row['Tutar'] < 0:
        return 'QB-K'
    elif row['Banka'] == 'TEB' and row['Fonksiyon Kodu 1'] == '701' and row['Tutar'] < 0:
        return 'TB-PI'
    elif row['Banka'] == 'TEB' and row['Fonksiyon Kodu 1'] == '701' and row['Tutar'] > 0:
        return 'TB-KI'
    elif row['Banka'] == 'TEB' and row['Fonksiyon Kodu 1'] != '701' and row['Tutar'] > 0:
        return 'TB-P'
    elif row['Banka'] == 'TEB' and row['Fonksiyon Kodu 1'] != '701' and row['Tutar'] < 0:
        return 'TB-K'
    else:
        return ''

Bulut_POS['KOD'] = Bulut_POS.apply(determine_kod, axis=1)

Bulut_POS['KOD1']= Bulut_POS['KOD'].str.split('-').str[1]

# Bulut_POS_Kodlu=Bulut_POS[["Hareket Tipi", "İşlem Kodu", "Banka", "Tutar", "Tarih", "Firma IBAN", "KOD", "KOD1"]].reset_index(drop=True)

POS_Toplam=Bulut_POS.groupby(['Tarih', 'Banka', "Firma IBAN", 'KOD',"KOD1"])['Tutar'].sum().reset_index()


POS_Ozet = POS_Toplam.pivot_table(index=['Banka', "Firma IBAN", 'Tarih'], columns='KOD1', values='Tutar', fill_value=0).reset_index()


POS_Ozet["Banka1"]= POS_Ozet["P"]+POS_Ozet["PI"]
POS_Ozet["Pos"]=-POS_Ozet["Banka1"]
POS_Ozet["Banka2"] =POS_Ozet["K"]+POS_Ozet["KI"]
POS_Ozet["Komisyon"]=-POS_Ozet["Banka2"]

POS_Ozet.reset_index(inplace=True)

melted_df = pd.melt(POS_Ozet, id_vars=["index", "Banka", "Firma IBAN", "Tarih"], value_vars=["Banka1", "Pos", "Banka2", "Komisyon"], 
                    var_name="Type", value_name="Tutar")
mapping = {'Banka1': 1, 'Pos': 2, 'Komisyon': 3, 'Banka2': 4}

melted_df["Sort"]= melted_df['Type'].map(mapping)
melted_df.sort_values(by=["index", "Sort"], inplace=True)
V3_Other=melted_df[melted_df["Tutar"]!=0].copy()
V3_Other.reset_index(drop=True)
V3_Other['Type'] = V3_Other['Type'].replace({'Banka1': 'Banka', 'Banka2': 'Banka'})

V3_Other.reset_index(inplace=True, drop=True)

IS_Pos=pd.read_excel(Config.iloc[5,1])
IS_Pos["Tarih"]=IS_Pos["Valör Tarihi"].dt.date

IS_Pos.sort_values(by=['Valör Tarihi'], inplace=True)
IS_Pos.reset_index(drop=True, inplace=True)

IS_Pos["Banka"]="İŞ BANKASI"
IS_Pos["Firma IBAN"]="TR950006400000143980027753"
IS_Pos["Pos"]=-IS_Pos["İşlem Tutarı"]
IS_Pos["Komisyon"]=IS_Pos["Komisyon Tutarı"]
IS_Pos["Banka1"]=IS_Pos["Net Tutar"]

IS_Pos.reset_index(inplace=True)
IS_Pos["index"]=IS_Pos["index"]+1000

melted_is = pd.melt(IS_Pos, id_vars=["index", "Banka", "Firma IBAN", "Tarih"], value_vars=["Banka1", "Pos", "Komisyon"], 
                    var_name="Type", value_name="Tutar")
mapping = {'Banka1': 2, 'Pos': 1, 'Komisyon': 3}

melted_is["Sort"]= melted_is['Type'].map(mapping)
melted_is.sort_values(by=["index", "Sort"], inplace=True)
V3_is=melted_is[melted_is["Tutar"]!=0].copy()
V3_is.reset_index(drop=True)
V3_is['Type'] = V3_is['Type'].replace({'Banka1': 'Banka'})

V3_is.reset_index(inplace=True, drop=True)

V3= pd.concat([V3_Other, V3_is]).reset_index(drop=True)

melted_Banka = pd.melt(Banka_Conf, id_vars=["Firma IBAN", "BANKA HESAP KODU"], value_vars=["Muhasebe Kodu", "Pos Kodu", "Pos Masraf Kodu"], 
                    var_name="Type", value_name="GLAccCode").dropna(subset=['GLAccCode'])

mapping_1 = {'Muhasebe Kodu': "Banka", 'Pos Kodu': 'Pos', "Pos Masraf Kodu": 'Komisyon'}

melted_Banka["Type"]= melted_Banka['Type'].map(mapping_1)

V3_data=V3.merge(melted_Banka, on=['Firma IBAN', 'Type'], how='left')

V3_Format=V3_data[["Tarih"]].copy()

V3_Format.rename(columns={"Tarih": 'DocumentDate'}, inplace=True)

V3_Format["DocumentNumber"]=V3_data["index"] + 1000
V3_Format["Description"]=V3_data["Banka"]+" POS İŞLEMLERİ"
V3_Format["BankCurrAccCode"] = np.where(V3_data["Type"] == "Banka", V3_data["BANKA HESAP KODU"], "")
V3_Format["BankOpTypeCode"]= 11
V3_Format["DueDate"]= ""
V3_Format["LineDescription"]= V3_data["Banka"]+" POS İŞLEMLERİ"
V3_Format["LineDocumentNumber"]= V3_data["Sort"]
V3_Format["DocumentTypeCode"]= ""
V3_Format["DocumentTypeDescription"]= ""
V3_Format["PaymentMethod"]= "Banka"
V3_Format["PaymentMethod"]= ""
V3_Format["GLAccCode"]= V3_data["GLAccCode"]
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

V3_Format.to_excel(Config.iloc[6,1], index=False)