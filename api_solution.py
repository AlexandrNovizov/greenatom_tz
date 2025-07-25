from pathlib import Path
import requests
import json

def main():

    request_data(date_start='2024-06-21', date_end='2025-06-21', curr="USD", dump_path='./data')
    request_data(date_start='2024-06-21', date_end='2025-06-21', curr="JPY", dump_path='./data')

    
def request_data(
        date_start: str,
        date_end: str, 
        curr: str, 
        limit: int = 100, 
        start: int = 0, 
        json_dump_folder: Path = None) -> list[dict]:

    total = float('inf')
    result = []

    while True:

        if total == start: break

        data = send_request(date_start, date_end, curr, limit, start)

        result.extend([
            {
                "rate": record["rate"],
                "time": record["tradetime"],
                "date": record["tradedate"]
            }
            for record in data[1]["securities"]
            if record["clearing"] == 'vk'
        ])

        pagination = data[1]['securities.cursor'][0]

        total = pagination['TOTAL']
        start += min(limit, total - start)

    if json_dump_folder != None:
        with open(f'{json_dump_folder}/{curr}.json', 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=4)

    return result

def send_request(date_start: str, date_end: str, curr: str, limit: int = 100, start: int = 0):
    params = {
        'lang':'ru',
        'iss.meta':'off',
        'iss.json':'extended',
        'iss.only':'history',
        'from': date_start,
        'till': date_end,
        'limit': limit,
        'sort_order':'DESC',
        'callback':'JSON_CALLBACK',
        'start': start
    }

    response = requests.get(
        url=f'https://iss.moex.com/iss/statistics/engines/futures/markets/indicativerates/securities/{curr}/RUB.json',params=params
    )

    if response.status_code == 200:
        return json.loads(response.text)
    else:
        raise RuntimeError(f"Error while request the API. Status code: {response.status_code}")



if __name__ == '__main__':
    main()