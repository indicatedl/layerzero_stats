import requests
from time import sleep
from sys import stderr, exit

from datetime import datetime
from art import text2art
from termcolor import colored
from alive_progress import alive_bar
import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
import inquirer
from inquirer.themes import load_theme_from_dict as loadth


file_wallets = 'files/wallets.txt'
file_table = 'files/LayerZeroStats.xlsx'


def worker(sort_type):
    wallets_data = {}
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    with alive_bar(len(wallets)) as bar:
        for address in wallets:
            if len(address) != 42:
                print(colored('\n\nЧел сюда надо адреса кошельков а не приватники!\n\n','light_yellow'))
                return
            while True:
                resp = requests.post('https://nftcopilot.com/p-api/layer-zero-rank/check', json={"address": address, "c": "check"})
                if resp.status_code not in (200, 201):
                    if 'Too Many Requests' not in resp.text:
                        print(resp.status_code, resp.text)
                    sleep(1)
                    continue
                break
            wallets_data[address] = resp.json()
            wallets_data[address]['rankUpdatedAt'] = datetime.utcfromtimestamp(int(wallets_data[address]['rankUpdatedAt'])/1000).strftime("%Y-%m-%d %H:%M:%S")
            bar()

    match sort_type:
        case "Сортировать кошельки по рангу (default)":
            wallets_data = dict(sorted(wallets_data.items(), key=lambda x: x[1]["rank"]))
        case "Сортировать кошельки по числу транзакций":
            wallets_data = dict(sorted(wallets_data.items(), key=lambda x: x[1]["txsCount"], reverse=True))
        case "Сортировать кошельки по обьемам":
            wallets_data = dict(sorted(wallets_data.items(), key=lambda x: x[1]["volume"], reverse=True))
        case "Сортировать кошельки по активным месяцам":
            wallets_data = dict(sorted(wallets_data.items(), key=lambda x: x[1]["distinctMonths"], reverse=True))
        case "Сортировать кошельки по сетям отправки":
            wallets_data = dict(sorted(wallets_data.items(), key=lambda x: x[1]["networks"], reverse=True))
        case "Сортировать кошельки по сетям назначения":
            wallets_data = dict(sorted(wallets_data.items(), key=lambda x: x[1]["destChains"], reverse=True))
        case "Не сортировать кошельки (будет как в исходном файле)":
            pass

    headers = ["Address", "Rank", "TxCount", "Volume", "Months", "SourceChains",
            "DestChains", "Contracts", "Top%TxCount", "Top%Volume", "Top%Months", 
            "Top%SourceChains", "Top%DestChains", "Top%Contracts", "Top%Final",
            "TotalLzUsers", "LastUpdate"]
    sheet.append(headers)

    for wallet, data in wallets_data.items():
        row = [wallet, data["rank"], data["txsCount"], data["volume"], data["distinctMonths"],
            data["networks"], data["destChains"], data["contracts"], f"{data['topInTxs']}%", 
            f"{data['topInVolume']}%",f"{data['topInUsageByMonth']}%", f"{data['topInUsageByNetwork']}%", 
            f"{data['topInDestChains']}%", f"{data['topInContracts']}%", f"{data['topFinal']}%",
            data["totalUsers"], data["rankUpdatedAt"]]
        sheet.append(row)

    for column in sheet.columns:
        for cell in column[1:]:
            cell.alignment = Alignment(horizontal='center')
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    sheet.column_dimensions['A'].width = 45  
    sheet.column_dimensions['D'].width = 10  
    sheet.column_dimensions['E'].width = 7.5 
    sheet.column_dimensions['F'].width = 12 
    sheet.column_dimensions['G'].width = 10 
    sheet.column_dimensions['H'].width = 10  
    sheet.column_dimensions['I'].width = 13  
    sheet.column_dimensions['J'].width = 12.5
    sheet.column_dimensions['K'].width = 13
    sheet.column_dimensions['L'].width = 17 
    sheet.column_dimensions['M'].width = 15 
    sheet.column_dimensions['N'].width = 14
    sheet.column_dimensions['O'].width = 10
    sheet.column_dimensions['P'].width = 12 
    sheet.column_dimensions['Q'].width = 17 

    workbook.save(file_table)
    print(colored(f'\nТаблица успешно составлена и сохранена в файл "{file_table}"!\n\n','light_yellow'))


def get_action() -> str:
    theme = {
        "Question": {
            "brackets_color": "bright_yellow"
        },
        "List": {
            "selection_color": "bright_blue"
        }
    }

    question = [
        inquirer.List(
            "action",
            message=colored("Выберите действие", 'light_yellow'),
            choices=["Получить статистику и составить Excel таблицу", 
                     "Выход"],
        )
    ]
    action = inquirer.prompt(question, theme=loadth(theme))['action']
    return action


def get_sort_type() -> str:
    theme = {
        "Question": {
            "brackets_color": "bright_yellow"
        },
        "List": {
            "selection_color": "bright_blue"
        }
    }

    question = [
        inquirer.List(
            "action",
            message=colored("Тип сортировки", 'light_yellow'),
            choices=["Сортировать кошельки по рангу (default)", 
                     "Сортировать кошельки по числу транзакций", 
                     "Сортировать кошельки по обьемам", 
                     "Сортировать кошельки по активным месяцам", 
                     "Сортировать кошельки по сетям отправки",
                     "Сортировать кошельки по сетям назначения",
                     "Не сортировать кошельки (будет как в исходном файле)"],
        )
    ]
    action = inquirer.prompt(question, theme=loadth(theme))['action']
    return action



def main() -> None:
    try:
        art = text2art(text="LAYERZERO  STATS", font="standart")
        print(colored(art,'light_blue'))
        art = text2art(text="from NFTCOPILOT.COM", font="cybermedum")
        print(colored(art,'light_cyan'))
        print(colored('Автор: t.me/cryptogovnozavod\n','light_cyan'))

        while True:
            action = get_action()

            match action:
                case 'Получить статистику и составить Excel таблицу':
                    sort_type = get_sort_type()
                    worker(sort_type)
                case 'Выход':
                    exit()
                case _:
                    pass
    except Exception as e:
        if 'Inappropriate ioctl for device' in str(e):
            print(colored('\n\nЗапустите программу в терминале! python3 main.py\n\n','red'))



if (__name__ == '__main__'):

    with open(file_wallets, 'r') as file:
        wallets = [row.split(":")[0].strip() if ":" in row else row.strip() for row in file]

    main()