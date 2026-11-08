import pandas as pd

파일들 = ["sk.csv", "samsung.csv", "samelec.csv", "hanjin.csv", "jyp.csv"]

표들 = []
for 파일 in 파일들:
    표 = pd.read_csv(파일)
    표들.append(표)

표_통합 = pd.concat(표들, ignore_index=True)

부채비율표 = 표_통합[표_통합["비율"] == "부채비율"]
정렬된표 = 부채비율표.sort_values("값")

print(정렬된표[["회사", "값"]].to_string(index=False))