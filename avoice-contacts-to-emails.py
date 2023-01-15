"""Personal email generation script

Dependencies: pandas, numpy, xpinyin (Use `pip install <package-name>`)

Instructions:
1. Export from Lark admin console “组织架构” “成员与部门” “批量导入或导出成员” “导出并修改成员信息”
2. Open exported file with Microsoft Excel, and save as "avoice-contacts.xlsx"
3. Run this script
4. Upload the output xlsx file "avoice-contacts2.xlsx" back to admin console
"""

import pandas as pd
import numpy as np
from xpinyin import Pinyin
import re

domain_name = "awdpi.org"

dt = pd.read_excel("avoice-contacts.xlsx", skiprows=0, header=1, converters={'手机号':str})
# dt = pd.read_csv("avoice-contacts.csv", skiprows=0, header=1)
# dt = dt[dt["帐号状态"]=="正常"]
# names = dt["姓名"]
# emails = dt["企业邮箱"]

p = Pinyin()

for i in range(len(dt)):
    # skip inactive accounts
    if dt.iloc[i]["帐号状态"] != "正常":
        continue
    # leave existing email addresses as they are
    if dt.iloc[i]["企业邮箱"] is not np.nan:
        continue
    # get the full name
    name = dt.iloc[i]["姓名"]
    name_stripped = name.replace(" ","").replace(".","")
    # all English alphabets, e.g. Rainie Wan
    if name_stripped.isalpha() and name_stripped.isascii():
        names = name.lower().split(" ")
        if len(names) == 1:
            email_addr = f"{names[0]}@{domain_name}"
        else:
            email_addr = f"{names[0]}.{names[-1]}@{domain_name}"
    else:
        # Chinese + English, e.g. 贺嘉文 Nancy
        if len(name.split(" ")) > 1:
            name = name.split(" ")[0]
        # all Chinese characters
        pinyins = p.get_pinyin(name, " ")
        pinyins = pinyins.lower()
        chinese_chars = pinyins.split(" ")
        email_first = "".join(chinese_chars[1:])
        email_last = chinese_chars[0] 
        email_addr = f"{email_first}.{email_last}@{domain_name}"
    # print(email_addr)
    assert re.match("[^@\s]+@[^@\s]+\.[a-zA-Z0-9]+$", email_addr)
    print(dt.iloc[i]["姓名"], email_addr)
    dt.at[i,"企业邮箱"] = email_addr

dt.to_excel("avoice-contacts2.xlsx", index=False)
