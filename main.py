from typing import Mapping

import pandas as pd

pd.set_option('display.width', None)
exclude_dates = pd.to_datetime([
    "2026-01-06",
    "2026-01-07",
    "2026-01-08",
    "2026-01-09",
    "2026-01-12",
    "2026-01-13",
    "2026-01-14",
    "2026-01-15",
    "2026-01-16",
    "2026-01-19",
    "2026-01-20",
    "2026-01-21",
    "2026-01-22",
    "2026-01-23",
    "2026-01-26",
    "2026-01-27",
    "2026-01-28",
    "2026-01-29",
    "2026-01-30"
])

print("Read File")
TransferRaw = pd.read_excel(r"SI-SO\Aug\Transfer.xlsx",dtype={"IMEI":str})
Transfer = TransferRaw.copy()
Order = pd.read_excel(r"SI-SO\Aug\Order.xlsx",dtype={"IMEI":str})
Client = pd.read_excel(r"SI-SO\Aug\Client Data - Final July.xlsx")
Info = pd.read_excel(r"SI-SO\Aug\IMEI Info.xlsx",dtype={"IMEI/SN":str})
Retail = pd.read_excel(r"SI-SO\Aug\Retail.xlsx",dtype={"IMEI":str})
PSM =pd.read_excel(r"SI-SO\Aug\PSM Final July.xlsx",sheet_name="Data Store")
PSMDirect = pd.read_excel(r"SI-SO\Aug\PSM Final July.xlsx",sheet_name="โค้ดรับตรง")
PSMPC = pd.read_excel(r"SI-SO\Aug\PSM Final July.xlsx",sheet_name="Data PC")

print("Read Completed")

Transfer = Transfer.sort_values(by=["transferTime"], ascending=[False])

Transfer = Transfer[
    (
        (
                (Transfer["fromDistributorName"] == "realme Brand Shop Sahathai Thungsong by INTERCOMPUTER AND IT")
                &
                (~pd.to_datetime(Transfer["invoiceDate"]).dt.normalize().isin(exclude_dates))
        )
        |
        (Transfer["fromDistributorCode"] == "DBBE5955")
        |
        (Transfer["IMEI"].isin([
            "867328080689950",
            "867328080871111",
            "863238080679394",
            "863861080527915"

        ]))
    )
]

# sort Product Category
Transfer = Transfer[
    ~Transfer["IMEI"].isin([
        "861079080884077"
    ])
]

Transfer = Transfer[
    (
            (Transfer["Product Category"] == "accessories") &
            (Transfer["productName"].isin([
                "realme Pad 3 5G",
                "realme Pad 3"
                ])
            )
            |
            (
                    Transfer["Product Category"] == "mobile")
    )
]
print("Transfer Completed")

#sort remark
Transfer["Remark Type"] = (
    Transfer["Remark"]
    .str.extract(r"([A-Za-z]+)$")
)
Transfer["Remark SBA"] = (
    Transfer["Remark"].str.split("-", n=2).str[1])
print("Remark1 Completed")

Remark = Transfer[
    (
    (Transfer["Remark Type"] == "Exchange") &
    (Transfer["fromDistributorCode"].isin([
        "DBBE5955"
    ])
    )
    |
    (
        (Transfer["Remark Type"] !="Exchange") &
        ~Transfer["Remark Type"].isin([
        "Demo",
        "Transfer",
        "EX",
        "Exchage",
        "Gift",
        "Gitf",
        "Return",
        "return"
        ])
    )
    )
]
print("Remark2 Completed")
SBA = Remark[
    (Remark["Remark SBA"] =="SBA")
]
print("Transfer Completed")

SBA["Date"] = SBA["transferTime"]

SBA.loc[
    SBA["Remark"].isin([
        "281-SBA-26030000001"
    ]),
    "Date"
] = "2026-02-27 00:00:00"
SBA.loc[
    SBA["Remark"].isin([
        "281-SBA-26040000001",
        "281-SBA-26040000002",
        "281-SBA-26040000004",
        "281-SBA-26040000010",
        "281-SBA-26040000013",
        "281-SBA-26040000014",
        "281-SBA-26040000009",
        "281-SBA-26040000008",
        "281-SBA-26040000007"
    ]),
    "Date"
] = "2026-03-31 00:00:00"
SBA.loc[
    SBA["IMEI"].isin([
        "861079081014435"
    ]),
    "Date"
] = "2026-06-26 00:00:00"
SBA.loc[
    SBA["IMEI"].isin([
        "860821081883355",
        "860821081883876",
        "860821081885012",
        "860821081885137",
        "860821081885814",
        "860821081886077",
        "860821081887158",
        "860821081888651",
        "860821081889857",
        "860821081895193",
        "860821081926576"
    ]),
    "Date"
]="2026-07-14 00:00:00"

Duplicate = SBA[
    SBA["IMEI"].duplicated(keep=False)
]

SBA =SBA.drop_duplicates(
    subset=["IMEI"],
      keep="first"
)

SBA.rename(
    columns={"toDistributorCode" : "Client Code"},
    inplace=True
)


SBA = SBA.merge(
    Client[
        [
            "Client Code",
            "Client Subtype",
            "Client Channel",
            "Main RSD",
            "Main Zone"
        ]
    ],
    left_on ="Client Code",
    right_on ="Client Code",
    how ="left"
)

print("Transfer Completed")

############################################################################################################

PSM.rename(
    columns={
        "Store ID" : "Store Code"},
    inplace=True
)
PSM["so_rsd"] = (
    PSM["RSD"].str.split("-", n=1).str[0])

PSM["so_rsm"] = (
    PSM["RSM"].str.split("-", n=1).str[0]
)
PSMDirect.rename(
    columns={"Store ID" : "Store Code"},
    inplace=True
)
############################################################################################################

Order = Order.sort_values(by=["Shipped Time"], ascending=[False])
Order = Order[
    Order["Warehouse Name"].isin([
        "Thailand Kerry Warehouse",
        "Thailand SF warehouse"
    ])
]

Order = Order[
    Order["Order Type"].isin([
        "Regular"
    ])
]
Order = Order[
    Order["Return Time"].isna()
]
Order.rename(
    columns={
        "Customer Code" : "Client Code",
        "Shipped Time" : "Date"
    },
    inplace=True
)
print("Order Completed")

###########################################################################
TransferSO = TransferRaw.sort_values(by=["transferTime"], ascending=[False])

#sort remark
TransferSO["Remark Type"] = (
    TransferSO["Remark"]
    .str.extract(r"([A-Za-z]+)$")
)
TransferSO["Remark SBA"] = (
    TransferSO["Remark"].str.split("-", n=2).str[1])
print("Remark1 Completed")

Remark = TransferSO[
    (
        TransferSO["IMEI"].isin([
            "867328080689950",
            "867328080871111",
            "863238080679394",
            "863861080527915"])
    )
    |
    (
        ~TransferSO["Remark Type"].isin([
        "Demo",
        "Transfer",
        "EX",
        "Exchage",
        "Gift",
        "Gitf",
        "Return",
        "return"
        ])
    )
]
print("Remark2 Completed")
SBASO = Remark[
    (Remark["Remark SBA"] =="SBA")
].copy()
print("TransferSO Completed")

SBASO = SBASO.drop_duplicates(
    subset=["IMEI"],
      keep="first"
)
SBASO.rename(
    columns={
        "toDistributorCode" : "Client Code",
        "transferTime" : "Date"
    },
    inplace=True
)
print("TransferSO Completed")
##################################################################################

Retail = Retail.sort_values(by=["Sale time"], ascending=[False])
Retail = Retail[
    Retail["Sale status"].isin([
        "Uploaded (already sold)",
        "In force"
    ])
]

Retail = Retail.drop_duplicates(
    subset=["IMEI"],
    keep="first"
)
print("Retail Completed")
##################################################################################

TransferSI = SBA[["IMEI","Client Code","Date"]]
OrderSI = Order[["IMEI","Client Code","Date"]]
OrderforSI = pd.concat(
    [TransferSI,OrderSI],
    ignore_index=True
)

TransferSO = SBASO[["IMEI","Client Code","Date"]]
OrderSO = Order[["IMEI","Client Code","Date"]]
OrderforSO = pd.concat(
    [TransferSO,OrderSO],
    ignore_index=True
)

##################################################################################

client_mapping = {
    "DBKH8368": "DBDJ2587",
    "DBHF7717": "DBCS7452",
    "DBAU4214": "DBSD0055",
    "DBBH7146": "DBUH6450",
    "DBBC9512": "DBZR6736",
    "DBQX6987": "DBBD1522",
    "DBFC8118": "DBBH6308",
    "DBHG3689": "DBGX2217"
}


for df in [OrderforSI, OrderforSO]:
    df["Client Code"] = (
        df["Client Code"]
        .replace(client_mapping)
    )

##################################################################################

for df in [OrderforSI, OrderforSO]:
    df["Date"] = pd.to_datetime(df["Date"])
    df.sort_values(by=["Date"],
                   ascending=False,
                   inplace=True
                   )

    df.drop_duplicates(subset = "IMEI",
                       keep="first",
                       inplace=True
                       )
    df.reset_index(drop=True, inplace=True)

##################################################################################

Info.rename(
    columns={"IMEI/SN" : "IMEI"},
    inplace=True
)
so_cond = Info.drop(
    columns=["Client Code"],
    errors="ignore"
).merge(
    OrderforSO[
        [
            "IMEI",
            "Client Code"
        ]
    ],
    on="IMEI",
    how="left"

)
Client_copy = Client.copy()

Client_copy["Client Code for Map"] = Client_copy["Client Code"]
Client_copy["Channel for Map"] = Client_copy["Client Channel"]

so_cond = so_cond.merge(
    Client_copy[
        [
            "Client Code",
            "Client Code for Map",
            "Client Name (GS-System)",
            "Client Type",
            "Client Subtype",
            "Client Channel",
            "Channel for Map",
            "Status",
            "Main Zone",
            "Main RSD"
        ]
    ],
    on="Client Code",
    how="left"
)

Channel_mapping = {
    "Hub" : "GT_Hub",
    "KD" : "GT_KD&Dealer",
    "Dealer" : "GT_KD&Dealer",
    "Brand shop_GT" : "GT_Brand shop_GT",
    "IT_CITY" : "KA",
    "JAYMART" : "KA",
    "Advice" : "KA",
    "TG" : "KA",
    "JIB" : "KA",
    "SINGER" : "KA",
    "MAXLINK" : "KA",
    "COM_7" : "Online_Com_7",
    "OPT_AIS" : "OPT_AIS",
    "OPT_TRUE" : "OPT-NOT AIS",
    "OPT_DTAC" : "OPT-NOT AIS",
    "OPT_CJ" : "OPT-NOT AIS",
    "After-Service" : "After-Service",
    "Others" : "Others",
    "Overseas" : "Overseas",
    "TH-realme" : "TH-realme"
}

so_cond["Channel for Map"] = so_cond["Channel for Map"].replace(Channel_mapping)

so_cond = so_cond.drop(
    columns=[
        "Sale time","Sale status"],
    errors="ignore"
)
so_cond = so_cond.merge(
    Retail[
        [
            "IMEI",
            "Sale time",
            "Sale status"
        ]
    ],
    on="IMEI",
    how="left"
)

so_cond = so_cond.merge(
    PSM[
        [
            "Store Code",
            "BU",
            "DB ID",
            "DB Name",
            "Channel Name",
            "Channel Responsible",
            "DB ID2",
            "DB Name2",
            "Province",
            "Area",
            "Area Code",
            "District",
            "Big Type",
            "Store Type",
            "Subclass2",
            "PC Total",
            "so_rsd",
            "so_rsm",
            "Zone"
        ]
    ],on="Store Code",
    how="left"
)

Client_DB2 = (
    Client[
        ["Client Code", "Client Channel"]
    ]
    .drop_duplicates("Client Code")
    .rename(columns={
        "Client Code": "DB ID2",
        "Client Channel": "DB2 channel"
    })
)

so_cond = so_cond.merge(
    Client_DB2,
    on="DB ID2",
    how="left"
)

so_cond = so_cond.merge(
    PSMDirect[
        [
            "Store Code",
            "Client ID สั่งตรง"
        ]
    ],on="Store Code",
    how="left"
)
so_cond.rename(
    columns={
        "Client ID สั่งตรง": "Client_Direct ID",
        "Client Name": "Client_Direct Name"
    },inplace=True
)


so_cond["SPU_Spec_Full"] = (
    so_cond["SPU Name"]
    .str.cat(so_cond["SKU Spec"], sep=" ")
)

so_cond["SPU_Spec_abbr"] = (
    so_cond["SPU_Spec_Full"]
    .str.extract(r"realme\s+([A-Za-z0-9])\s+(\d+)GB\s+(\d+)GB")
    .apply(lambda x:f"{x[0]}_{x[2]}+{x[1]}",axis=1)
)
so_cond["spec_tier"] = None
so_cond["E_date"] = so_cond[["Activate Time","Sale time"]].min(axis=1)
so_cond["ST Client"] = None
so_cond["ST Client Name"] = None
so_cond["ST Client Channel"] = None
so_cond["Store id by client"] = None
so_cond["Si_wh"] = None
so_cond["Store No Pc 15 Days"] = None

so_cond = so_cond[
    [
        "IMEI",
        "Order Type",
        "Product Type",
        "SKU ID",
        "Series Name",
        "spec_tier",
        "SPU Name",
        "SKU Spec",
        "SPU_Spec_Full",
        "SPU_Spec_abbr",
        "Activate Time",
        "Sale time",
        "E_date",
        "Sale status",
        "Si_wh",
        "Client Code",
        "Client Name (GS-System)",
        "Client Type",
        "Client Channel",
        "Channel for Map",
        "Status",
        "Main Zone",
        "Main RSD",
        "ST Client",
        "ST Client Name",
        "ST Client Channel",
        "Store Code",
        "Store id by client",
        "Store Name",
        "BU",
        "DB ID",
        "DB Name",
        "Channel Name",
        "Channel Responsible",
        "DB ID2",
        "DB Name2",
        "DB2 channel",
        "Province",
        "Area",
        "Area Code",
        "District",
        "Big Type",
        "Store Type",
        "Subclass2",
        "Client_Direct ID",
        "Client_Direct Name",
        "PC Total",
        "Store No Pc 15 Days",
        "Zone",
        "so_rsd",
        "so_rsm"
    ]
]
so_condforCash = so_cond[
    [
        "IMEI",
        "Order Type",
        "Product Type",
        "SKU ID",
        "Series Name",
        "spec_tier",
        "SPU Name",
        "SKU Spec",
        "SPU_Spec_Full",
        "SPU_Spec_abbr",
        "Activate Time",
        "Sale time",
        "E_date",
        "Client Code",
        "Client Code for Map",
        "Client Name (GS-System)",
        "Client Type",
        "Client Subtype",
        "Client Channel",
        "Channel for Map",
        "Status",
        "Main Zone",
        "Main RSD",
        "ST Client",
        "ST Client Name",
        "ST Client Channel",
        "Store Code",
        "Store id by client",
        "Sale status"
    ]
]
##################################################################################

print("so_cond Completed")

##################################################################################

with (pd.ExcelWriter(r"SI-SO\Aug\Result.xlsx",engine="openpyxl") as writer):
    SBA.to_excel(writer,sheet_name="Transfer",index=False)
    Order.to_excel(writer,sheet_name="Order",index=False)
    SBASO.to_excel(writer, sheet_name="TransferSO", index=False)
    Retail.to_excel(writer, sheet_name="Retail", index=False)
    OrderforSI.to_excel(writer, sheet_name="OrderforSI", index=False)
    OrderforSO.to_excel(writer, sheet_name="OrderforSO", index=False)
    so_cond.to_excel(writer, sheet_name="so_cond", index=False)
    so_condforCash.to_excel(writer, sheet_name="so_condforCash", index=False)

print("Run Completed")
