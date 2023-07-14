import os
import sys
import urllib.request
import openpyxl
import json

# Excel File Library
start, num = 1, 0

excel_file = openpyxl.Workbook()
excel_sheet = excel_file.active
excel_sheet.column_dimensions["B"].width = 100
excel_sheet.column_dimensions["C"].width = 100
excel_sheet.append(["랭킹", "제목", "최저가", "최고가", "몰이름", "링크"])


client_id = "h5dpSfKrVTjaKh0ckcUw"
client_secret = "ThoKgI5GGT"

for index in range(10):
    start_number = start + (index * 10)
    encText = urllib.parse.quote("BG0MTN703PI")
    url = (
        "https://openapi.naver.com/v1/search/shop.json?query="
        + encText
        + "&display=100&start="
        + str(start_number)
        + "&sort=asc"
    )  # JSON 결과
    # url = "https://openapi.naver.com/v1/search/blog.xml?query=" + encText # XML 결과
    request = urllib.request.Request(url)
    request.add_header("X-Naver-Client-Id", client_id)
    request.add_header("X-Naver-Client-Secret", client_secret)
    response = urllib.request.urlopen(request)
    rescode = response.getcode()
    if rescode == 200:
        # response_body = response.read()
        # data = response_body.decode("urf-8")
        data = json.loads(response.read().decode("utf-8"))
        for item in data["items"]:
            num += 1
            excel_sheet.append(
                [
                    num,
                    item["title"],
                    item["lprice"],
                    item["hprice"],
                    item["mallName"],
                    item["link"],
                ]
            )

        # print(response_body.decode("utf-8"))
    else:
        print("Error Code:" + rescode)

excel_file.save("IT.xlsx")
excel_file.close()
