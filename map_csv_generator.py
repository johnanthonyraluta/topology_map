import pandas as pd
import sys

hostname = ["PLDT CAPAS",
            "PLDT LUISITA INDUSTRIAL PARK",
            "PLDT MANAOAG",
            "PLDT PANIQUI",
            "PLDT PANTAL",
            "PLDT TARLAC",
            "SMART 4943 CARMEN",
            "SMART 587 TARLAC 2 SAN NICOLAS",
            "SMART C03 SAN CARLOS",
            "SMART C05 MANGALDAN",
            "SMART C07 ALAMINOS",
            "SMART C10 ASINGAN",
            "SMART C40 MANGATAREM",
            "SMART C43 MANAOAG",
            "SMART N26 TIBAG",
            "SMART O39 CONCEPCION TP",
            "SMART Y30 LINGAYEN 3 - DOMALANDAN",
            "SMART Y31 BURGOS",
            "SMART C45 BINALONAN",
            "SMART 5376 TAMARO"]
for host in hostname:
    path = r"data_output\\"
    host_path = path + host + ".xlsx"
    writer = pd.ExcelWriter(host_path, engine = 'xlsxwriter')
    df = pd.read_excel("src\\PLDT TNT PH2 - Port Mapping, IP Addressing and Device Information Draft.xlsx",sheet_name="port_mapping")
    #print(df.head(100))
    source_target_df = df[df.SOURCE_SITE.str.contains(host)]
    #print(source_target_df.head(100))
    target_source_df = source_target_df.rename(columns={"HOSTNAME":"TARGET_HOSTNAME","SOURCE_SITE":"TARGET_SITE",\
                        "SOURCE_DEVICE_TYPE":"TARGET_DEVICE_TYPE","SOURCE_REGION":"TARGET_REGION","TARGET_HOSTNAME":\
                        "HOSTNAME","TARGET_SITE":"SOURCE_SITE","TARGET_DEVICE_TYPE":"SOURCE_DEVICE_TYPE",\
                        "TARGET_REGION":"SOURCE_REGION"})
    new_df = pd.concat([source_target_df,target_source_df],ignore_index=True)

    new_df[["SOURCE_INTERFACE_ID"]] = ""
    new_df[["SLOT"]] = ""
    new_df[["SOURCE_BUNDLE_INTERFACE"]] = ""
    new_df[["SOURCE_IPV4_ADDRESS"]] = ""
    new_df[["SOURCE_IPV6_ADDRESS"]] = ""
    new_df[["TARGET_INTERFACE_ID"]] = ""
    new_df[["Unnamed: 16"]] = ""
    new_df[["TARGET_BUNDLE_INTERFACE"]] = ""
    new_df[["TARGET_IPV4_ADDRESS"]] = ""
    new_df[["TARGET_IPV6_ADDRESS"]] = ""
    new_df2 = new_df[["SOURCE_SITE","SOURCE_DEVICE_TYPE","HOSTNAME","SOURCE_REGION"]].copy()
    new_df2 = new_df2.rename(columns={"SOURCE_SITE":"SITE_NAME","SOURCE_DEVICE_TYPE":"DEVICE_TYPE",\
                                       "SOURCE_REGION":"REGION"})
    new_df2["KEY"] = new_df2["HOSTNAME"]

    print(new_df2.head(100))
    #new_df.to_csv("data_output/binmaley.csv", index=False)
    pd.DataFrame(new_df2).to_excel(writer, sheet_name="device_information", index=False)
    pd.DataFrame(new_df).to_excel(writer, sheet_name="port_mapping", index=False)
    writer.save()
print("Done...")