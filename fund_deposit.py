import glob
import subprocess
import os
import pandas as pd

def main():
    testcasesize = 8
    path = 'C:\eBanking_API_Testing\*\\testcase_result_*.xlsx'
    filename = 'output.xlsx'
    pd_output = pd.read_excel('LOG1826 Timeout issue (2022-07-11) 11 sequential testing for FD.xlsx')
    summary = pd.read_excel('LOG1826 Timeout issue (2022-07-11) 11 sequential testing for FD.xlsx', sheet_name=1)
    profilelist=[]
    for i in range(testcasesize):
        for j in range(36):
            profilelist.append(i)
    print(pd_output)

    for i,file in enumerate(sorted(glob.glob(path), key = len)):
        print(file)
        if i == testcasesize:
            break
        pd_input = pd.read_excel(file, sheet_name = "input", header = 1)
        pd_result = pd.read_excel(file, sheet_name = "result", header = 1)


        print(pd_input)
        print(pd_result)
        print(pd_result.dtypes)
        pd_result.columns = pd_result.columns.str.strip()
        pd_merge = pd.merge(pd_input,pd_result,on=['Test_Case','Fixed_Deposit_Placement'])
        pd_output = pd.concat([pd_output,pd_merge])

    print(pd_output.dtypes)
    pd_output.insert(0, "Profile", profilelist, True)
    with pd.ExcelWriter(filename) as writer:
        pd_output.to_excel(writer,"Result", index=False)
        summary.to_excel(writer,"Finding", index=False)

    os.system("start " + filename)

if __name__ == '__main__':
    main()