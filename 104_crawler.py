import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import threading
import time

# version: 2024/09/03
def loading_message(stop_event):
    while not stop_event.is_set():
        print("Loading",end='')
        for i in range(5):
            print(".",end='',flush=True)
            time.sleep(1)
        print("\r",end='')

def get_job_data(keyword,location_numbers):
    # 多啟用一線程來執行Loading
    stop_event = threading.Event()
    loading_thread = threading.Thread(target=loading_message, args=(stop_event,))
    loading_thread.start()

    url = f'https://www.104.com.tw/jobs/search/?jobsource=index_s&keyword={keyword}&area={location_numbers}&mode=s&page=1'
    headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36"
    }
    res = requests.get(url, headers=headers)
    soup = BeautifulSoup(res.text, "html.parser")
    page = 1

    job_list = []

    while soup.find_all('article', class_="b-block--top-bord job-list-item b-clearfix js-job-item"):
        for job in soup.find_all('article', class_="b-block--top-bord job-list-item b-clearfix js-job-item"):
            job_data = {}
            job_data["職缺名稱"] = job['data-job-name']  # 職缺名稱
            job_data["職缺連結"] = 'https:' + job.a['href']  # 職缺連結
            ul = job.find("ul",{"class": "b-list-inline b-clearfix job-list-intro b-content"})
            li = ul.find_all("li")
            job_data["經歷要求"] = li[1].text  
            job_data["公司名稱"] = job['data-cust-name']  # 公司名稱
            job_data["工作地區"] = job.select('ul.b-list-inline.b-clearfix.job-list-intro.b-content li')[0].text  # 工作地區
            
            salary_info = job.find('div', class_="job-list-tag b-content")

            # 處理薪資資訊
            if salary_info.select('span') and salary_info.span.text == "待遇面議":
                e = salary_info.span.text
            else:
                e = salary_info.a.text
            
            job_data["薪資待遇"] = e
            job_data["計薪方式"] = e[:2]  # 計薪方式
            salary = ''
            for char in e:
                if char.isdigit() or char == '~':
                    salary += char
            
            if '~' in salary:
                low_salary = salary[:salary.find('~')]
                high_salary = salary[salary.find('~') + 1:]
            else:
                low_salary = salary
                high_salary = ''
            
            job_data["薪資下限"] = int(low_salary) if low_salary.isdigit() else 40000
            job_data["薪資上限"] = int(high_salary) if high_salary.isdigit() else 40000
            

            job_list.append(job_data)
        
        page += 1
        res = requests.get(f'https://www.104.com.tw/jobs/search/?ro=0&kwop=7&keyword={keyword}&expansionType=area%2Cspec%2Ccom%2Cjob%2Cwf%2Cwktm&area={location_numbers}&order=14&asc=0&page={page}&mode=s&jobsource=index_s&langFlag=0&langStatus=0&recommendJob=1&hotJob=1')
        soup = BeautifulSoup(res.text, "html.parser")
    # 設置停止loading顯示
    stop_event.set()
    loading_thread.join()

    return job_list

def analyze_job_data(job_list):
    df = pd.DataFrame(job_list)
    total_jobs = len(df)
    max_salary = df["薪資上限"].max()
    min_salary = df["薪資下限"].min()
    avg_salary = df[["薪資下限", "薪資上限"]].mean().mean()

    return total_jobs, max_salary, min_salary, avg_salary

def save_to_excel(df,filename,keyword):
    if not os.path.exists(filename):
        wb = Workbook()
        ws = wb.active
        ws.title = f"104 {keyword}的職缺"
        for r in dataframe_to_rows(df,index=False,header=True):
            ws.append(r)
        wb.save(filename)
    else:
        sheet_name = f"104 {keyword}的職缺"
        with pd.ExcelWriter(filename, mode='a',if_sheet_exists="new") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

def get_location_numbers(location_list):
    location_number_list = []
    for i in location_list:
        # if i in location_dict.keys:
        #     location_number_list.append(location_dict[i])
        for key in location_dict.keys():
            if len(set(i) & set(key)) >= 3:
                location_number_list.append(location_dict[key])
    location_numbers = '%2C'.join(location_number_list)
    return location_numbers

location_dict = {
    '台北市': '6001001000',
    '新北市': '6001002000',
    '宜蘭縣': '6001003000',
    '基隆市': '6001004000',
    '桃園市': '6001005000',
    '新竹縣市': '6001006000',
    '苗栗縣': '6001007000',
    '台中市': '6001008000',
    '彰化縣': '6001010000',
    '南投縣': '6001011000',
    '雲林縣': '6001012000',
    '嘉義縣市': '6001013000',
    '台南市': '6001014000',
    '高雄市': '6001016000',
    '屏東縣': '6001018000',
    '台東縣': '6001019000',
    '花蓮縣': '6001020000',
    '澎湖縣': '6001021000',
    '金門縣': '6001022000',
    '連江縣': '6001023000',
}

if __name__ == "__main__":
    keyword = input("請輸入要搜尋的職業關鍵字:")
    location = input("請輸入要搜尋的地區(ex:台北市 台中市，請用空格隔開):")
    location_list = location.split()
    location_numbers = get_location_numbers(location_list)
    job_list = get_job_data(keyword,location_numbers=location_numbers)
    total_jobs, max_salary, min_salary, avg_salary = analyze_job_data(job_list)
    
    print("資料爬取完成")
    
    user_input = input("請輸入輸出選項:(1.印出資料, 2.存成excel):")
    current_date = datetime.now().strftime("%Y-%m-%d")

    if user_input == '1' or user_input == '1.':
        print(f"執行日期: {current_date}")
        print(f"職缺數量: {total_jobs}")
        print(f"薪資平均:{avg_salary:.2f}")
        print(f"最高薪資:{max_salary}")
        print(f"最低薪資:{min_salary}")
    elif user_input == '2' or user_input == '2.':
        os.makedirs("vacancies_excel",exist_ok=True)
        df = pd.DataFrame(job_list)
        filename = f"vacancies_excel/{current_date}_職缺資料.xlsx"
        save_to_excel(df, filename, keyword)
        print(f"資料已存入 {filename}")
    else:
        print("無效選項")