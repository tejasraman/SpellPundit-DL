import requests
from datetime import datetime
from tqdm import tqdm
from time import sleep
import xlsxwriter
oper = "vgrade_module"
default_auth = "_gid=GA1.2.1716747121.1685722542; PHPSESSID=jius3llf0t7mbdrpjuehqm33e4; _ga_9ET3S5JW2T=GS1.1.1685722542.1.1.1685722789.0.0.0; _ga=GA1.2.1369650560.1685722542; _gat_gtag_UA_113244794_1=1"
auth = input("Please enter auth token (say \"def\" to set to default): ")

def main(auth):
    pid = input(
        "Please enter the Pack ID. Item IDs are appended to end: ")
    mid = input("Please enter your Module ID: ")
    packnum = input("Please enter the number of items in the pack: ")
    filename = input("Please enter the name of the file you want to save to: ")

    if auth.lower() == "def":
        auth = default_auth

    def request(num, rpnum, pid, auth):
        headers = {
            'Cookie': auth,
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.42',
            'sec-ch-ua-platform': 'Windows',
            'Accept': '*/*',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'Content-Length': '125',
            'Content-Type': 'application/x-www-form-urlencoded',
            'DNT': '1',
            'Host': 'spellpundit.com',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'sec-ch-ua-mobile': '?0'
        }

        now = datetime.now()

        data = {
            'mode': oper,
            'flag': 's',
            'bt': 'n',
            'vocab_cnt': num,
            'vocab_id': f"{pid}+{rpnum}",
            'set_id': pid,
            'module_id': int(mid),
            'fav_ind': '0',
            'start_date': now.strftime("%Y-%m-%d %H:%M:(%S - 1)")
        }

        data = requests.post(
            'https://spellpundit.com/spell/index.php', headers=headers, data=data).text

        return data

    lst = []
    print("Step 1: Fetch Words")
    for i in tqdm(range(1, (int(packnum) + 1)), colour="#00dff1"):
        lst.append(request(i, f"0 + {i + 1}", pid, auth))
    print("Step 2: Save to Excel file")
    workbook = xlsxwriter.Workbook(f'{filename}.xlsx')
    worksheet = workbook.add_worksheet()

    for i, j in enumerate(tqdm(lst, colour="#ffd100")):
        worksheet.write(f'A{i+1}', j)

    workbook.close()


main(auth)

while True:
    inp = input("Run again? (yes / no) > ")
    if inp.lower() == "yes":
        main(auth)
    else:
        break
