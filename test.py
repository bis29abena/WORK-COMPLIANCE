import pandas as pd
from collections import Counter

idxs = []
REGION_FILES = ["ASHANTI_REGION.xlsx", "BONO_AHAFO_REGION.xlsx", "KUMASI.xlsx", "NORTHERN_REGION.xlsx",
                "UPPER_EAST_REGION.xlsx", "UPPER_WEST_REGION.xlsx", "WESTERN_REGION.xlsx"]
# open the necessary excel files

for region in REGION_FILES:
    try:
        region_sheet2 = pd.read_excel(io=region, sheet_name="Sheet2")
        region_sheet1 = pd.read_excel(io=region, sheet_name="Sheet1")
    except FileNotFoundError:
        print(FileNotFoundError)
    else:
        print(f"{region} Files Opened Sucessfully")

        employers_list = region_sheet2.query('NUMBER_OF_EMPLOYEES > 1')["FEMPNAME WITHOUT DUPLICATES"].tolist()

        for employer in employers_list:
            idx = region_sheet1.index[region_sheet1["FEMPNAME"] == employer].tolist()
            idxs.extend(idx)

        more_employees = region_sheet1.loc[idxs, ["FEMPNAME", "EMP_TOWN", "EMP_ACTUAL_LOCATION"]]
        more_employees = more_employees.sort_values(by="FEMPNAME")
        more_employees = more_employees.drop_duplicates(subset="FEMPNAME")

        towns_in_list = pd.DataFrame(list(Counter(more_employees["EMP_TOWN"].tolist()).items()), columns=["TOWN_NAME",
                                                                                                          "TOTAL"])
        towns_in_list = towns_in_list.sort_values(by="TOTAL", ascending=False)

        with pd.ExcelWriter(region, mode="a") as writer:
            more_employees.to_excel(writer, sheet_name="Sheet3")
            towns_in_list.to_excel(writer, sheet_name="Sheet4")
        idxs = []

        print(f"{region} saving file done")
