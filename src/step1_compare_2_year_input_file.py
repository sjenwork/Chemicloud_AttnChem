#filecmp.cmp 函數是純粹針對兩個檔案內容進行比較，只要實際的內容相同，則判定為相同。https://officeguide.cc/python-how-to-compare-two-files/
import filecmp
import glob
import os
import pandas as pd
import shutil
#setting
year_1 = "2021"
year_2 = "2022"
allFiles_1 = [os.path.basename(x) for x in glob.glob(f"../{year_1}_input/*")]
allFiles_1 = [x.lower() for x in allFiles_1]#統一 小寫附檔名
allFiles_2 = [os.path.basename(x) for x in glob.glob(f"../{year_2}_input/*")]
allFiles_2 = [x.lower() for x in allFiles_2]#統一 小寫附檔名
#交集檔案 (intersection)
set_1 = set(allFiles_1)
set_2 = set(allFiles_2)
set_3 = set_1 & set_2
allFiles_intersection = list(set_3)
allFiles_intersection.sort()
print(allFiles_intersection)

# 檢查 file1.txt 與 file2.txt 是否相同
for file_name in allFiles_intersection:
    if filecmp.cmp(f"../{year_1}_input/{file_name}", f"../{year_2}_input/{file_name}"):
        print(f"{year_1}, {year_2} {file_name} 檔案相同")

        
        if file_name =="65_67.xlsx":
            shutil.copyfile(f"../{year_1}_processed/{file_name}", f"../{year_2}_processed/{file_name}")

        try:
            file_name = file_name.replace("_2","").split(".")[0]+".xlsx"
            print(file_name)
            df = pd.read_excel(f"../{year_1}_processed/{file_name}")
            df.to_excel(f"../{year_2}_processed/{file_name}", index=False)

        except Exception as e:
            print(e)
            if e == f"[Errno 2] No such file or directory: '../{year_1}_processed/53_55.xlsx'":

                df = pd.read_excel(f"../{year_1}_processed/53.xlsx")
                df.to_excel(f"../{year_2}_processed/53.xlsx", index=False)
                df = pd.read_excel(f"../{year_1}_processed/54.xlsx")
                df.to_excel(f"../{year_2}_processed/54.xlsx", index=False)
                df = pd.read_excel(f"../{year_1}_processed/55.xlsx")
                df.to_excel(f"../{year_2}_processed/55.xlsx", index=False)
                # shutil.copyfile(f"../{year_1}_processed/53.xlsx", f"../{year_2}_processed/53.xlsx")
                # shutil.copyfile(f"../{year_1}_processed/54.xlsx", f"../{year_2}_processed/54.xlsx")
                # shutil.copyfile(f"../{year_1}_processed/55.xlsx", f"../{year_2}_processed/55.xlsx")

                continue
        # try:
        #     file_name = file_name.replace("_2","").split(".")[0]+".xlsx"
        #     print(file_name)
        #     shutil.copyfile(f"../{year_1}_processed/{file_name}", f"{file_name}")

        # except Exception as e:
        #     print(e)
        #     if e == f"[Errno 2] No such file or directory: '../{year_1}_processed/53_55.xlsx'":
        #         shutil.copyfile(f"../{year_1}_processed/53.xlsx", f"53.xlsx")
        #         shutil.copyfile(f"../{year_1}_processed/54.xlsx", f"54.xlsx")
        #         shutil.copyfile(f"../{year_1}_processed/55.xlsx", f"55.xlsx")

        #         continue

    else:
        print(f"{year_1}, {year_2} {file_name} 不同")


