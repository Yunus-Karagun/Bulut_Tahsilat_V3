import numpy as np
import pandas as pd

Config=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Config")
SGK_Conf=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="SGK")
Banka_Conf=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="Banka")
SGK_Conf=pd.read_excel(r"D:\BulutTahsilat\Config.xlsx", sheet_name="SGK")
SGK_Ekstre=pd.read_excel(Config.iloc[11,1])

Bulut_All=pd.read_excel(Config.iloc[4,1])
Bulut_All["Tarih"]=Bulut_All["Tarih"].dt.date

SGK_Bordro=SGK_Ekstre[SGK_Ekstre["Belge\nTipi\nAçıklaması"]=="Ücret Bordrosu İcmali"][["Hesap\nKodu", "Hesap\nAçıklaması", "Fiş\nTarihi", "Gönderen\nOfis\nKodu", 'Maliyet\nMerkezi', 'Bakiye\n(Y)']]

Ekstre = SGK_Bordro.groupby(['Fiş\nTarihi', 'Hesap\nKodu', 'Hesap\nAçıklaması'])['Bakiye\n(Y)'].sum().reset_index()

Ekstre = Ekstre.rename(columns={
    'Fiş\nTarihi': 'Tarih',
    'Hesap\nKodu': 'GLAccCode',
    'Hesap\nAçıklaması': 'Hesap_Aciklamasi',
    'Bakiye\n(Y)': 'Tutar'
})

Bulut_SGK=Bulut_All[(Bulut_All["Fonksiyon Kodu 2"]=="TAHFATURATAH_SGKMOSIP")&(Bulut_All["Tarih"]==Config.iloc[10,1].date())].copy()

Bulut_SGK['SGK İşyeri Sicil No'] = Bulut_SGK['Açıklama'].str.extract(r"Sicil No: (\d{26})")

SGK_Kod = pd.melt(SGK_Conf, id_vars=["MK", "SGK İşyeri Sicil No", "Mağaza Adı"], value_vars=["ÖDENECEK SOSYAL GÜVENLİK KESİNTİLERİ", "ÖDENECEK İŞSİZLİK PRİMLERİ", "ÖDENECEK SGDP PRİMLERİ"], 
                    var_name="Type", value_name="GLAccCode").sort_values(by="SGK İşyeri Sicil No").dropna(subset=['GLAccCode']).reset_index(drop=True)

Ekstre_Kodlu=Ekstre.merge(SGK_Kod, how="left", on="GLAccCode")

Ekstre_İsyeri_Toplam=Ekstre_Kodlu.groupby(['SGK İşyeri Sicil No', 'Mağaza Adı'])['Tutar'].sum().reset_index()
Ekstre_İsyeri_Toplam.rename(columns={'Tutar': 'Tahakkuk'}, inplace=True)

Odeme_Toplam=Bulut_SGK.groupby(['SGK İşyeri Sicil No'])['Tutar'].sum().reset_index()

Odeme_Toplam.rename(columns={'Tutar': 'Banka1'}, inplace=True)

SGK_Ozet=Odeme_Toplam.merge(Ekstre_İsyeri_Toplam, how="left", on="SGK İşyeri Sicil No")

SGK_Ozet["Gelir"]=SGK_Ozet["Tahakkuk"]-SGK_Ozet["Banka1"]

Data=Bulut_SGK[["İşlem Kodu", "Banka", "Tarih", "Firma IBAN", "Açıklama", "SGK İşyeri Sicil No", "Tutar"]].reset_index(drop="True")

SGK_data=Data.merge(SGK_Ozet[["SGK İşyeri Sicil No", "Mağaza Adı", "Gelir"]], on="SGK İşyeri Sicil No", how="left")

SGK_data["Açıklama"]="SGK Ödemesi-"+SGK_data["Banka"]+"-"+SGK_data["SGK İşyeri Sicil No"]+"-"+SGK_data["Mağaza Adı"]

melted_df = pd.melt(SGK_data, id_vars=["Tarih", "İşlem Kodu", "Firma IBAN", "Banka", "Açıklama", "SGK İşyeri Sicil No", "Mağaza Adı"], value_vars=[ "Tutar", "Gelir"], 
                    var_name="Type", value_name="Tutar1").sort_values(by="İşlem Kodu").reset_index(drop=True)

melted_df['Type'] = melted_df['Type'].replace({'Tutar': 'Banka'})

V3_data_1=melted_df.merge(Banka_Conf[["Firma IBAN", "BANKA HESAP KODU", "Muhasebe Kodu"]], on="Firma IBAN", how="left")

Ekstre_data=Ekstre_Kodlu[["SGK İşyeri Sicil No"]].copy()

Ekstre_data["Tutar1"]=-Ekstre_Kodlu[["Tutar"]]
Ekstre_data["Muhasebe Kodu"]=Ekstre_Kodlu[["GLAccCode"]]

V3_data = pd.concat([V3_data_1, Ekstre_data]).sort_values(by=["SGK İşyeri Sicil No", "Muhasebe Kodu"]).reset_index(drop=True)

# Setting the option to opt-in to the future behavior
pd.set_option('future.no_silent_downcasting', True)

# Forward fill the specific columns
V3_data[["Tarih", 'İşlem Kodu', "Açıklama"]] = V3_data[["Tarih", 'İşlem Kodu', "Açıklama"]].ffill()

# Explicitly infer the objects dtype
V3_data = V3_data.infer_objects(copy=False)

V3_data.drop(V3_data[V3_data['Tutar1'] == 0].index, inplace=True)

V3_data.reset_index(drop=True, inplace=True)

V3_data['LineDocumentNumber'] = V3_data.groupby('İşlem Kodu').cumcount() + 1

V3_data.rename(columns={'Tutar1': 'Tutar'}, inplace=True)

V3_Format=V3_data[["Tarih"]].copy()
V3_Format.rename(columns={"Tarih": 'DocumentDate'}, inplace=True)

V3_Format["DocumentNumber"]=V3_data["İşlem Kodu"]
V3_Format["Description"]=V3_data["Açıklama"]
V3_Format["BankCurrAccCode"] = np.where(V3_data["Type"] == "Banka", V3_data["BANKA HESAP KODU"], "")
V3_Format["BankOpTypeCode"]= 25
V3_Format["DueDate"]= ""
V3_Format["LineDescription"]= V3_data["Açıklama"]
V3_Format["LineDocumentNumber"]= V3_data["LineDocumentNumber"]
V3_Format["DocumentTypeCode"]= ""
V3_Format["DocumentTypeDescription"]= ""
V3_Format["PaymentMethod"]= "Banka"
V3_Format["GLAccCode"]= np.where(V3_data["Type"] == "Gelir", "602.02.01.002", V3_data["Muhasebe Kodu"])
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

V3_Format.to_excel(Config.iloc[8,1], index=False)