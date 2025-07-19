import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os

import pandas as pd
import requests

load_dotenv()

currency_pair_left = input('Введите левую пару(USD/RUB) ') or 'USD/RUB'
currency_pair_right = input('Введите левую пару(JPY/RUB)  ') or 'JPY/RUB'
from_date_input = str(input('Введите дату начала(2025-06-01) ')) or '2025-06-01'
to_date_input = str(input('Введите дату окончания(2025-06-30) ')) or '2025-06-30'


def fetch_currency_data(currency_pair: str, from_date: str, to_date: str) -> dict:
    url = f'https://iss.moex.com/iss/statistics/engines/futures/markets/indicativerates/securities/{currency_pair}.json?lang=ru&from={from_date}&till={to_date}&iss.meta=off&iss.json=extended&iss.meta=off&limit=100&start=0&sort_order=DESC&iss.meta=off&iss.json=extended&callback=JSON_CALLBACK&lang=ru'
    response = requests.get(url).json()
    data = {
        'status': 1,
        'from_date': from_date,
        'to_date': to_date,
        'currency_pair': currency_pair,
        'data': []
    }
    for i in response[1:]:
        data['data'] = i.get('securities')
    if data['data']:
        return data
    else:
        raise KeyError('Произошла ошибка получения данных, проверьте правильность введенных атрибутов')


def create_excel(usd_data: dict, jpy_data: dict) -> int:
    df_jpy = pd.DataFrame(jpy_data['data'])
    df_usd = pd.DataFrame(usd_data['data'])

    df_jpy_main = df_jpy[df_jpy['clearing'] == 'pk'].copy()
    df_usd_main = df_usd[df_usd['clearing'] == 'pk'].copy()

    df_jpy_main.rename(columns={
        'tradedate': f'Дата {jpy_data["currency_pair"]}',
        'rate': f'Курс {jpy_data["currency_pair"]}',
        'tradetime': f'Время {jpy_data["currency_pair"]}'
    }, inplace=True)

    df_usd_main.rename(columns={
        'tradedate': f'Дата {usd_data["currency_pair"]}',
        'rate': f'Курс {usd_data["currency_pair"]}',
        'tradetime': f'Время {usd_data["currency_pair"]}'
    }, inplace=True)

    result_df = pd.merge(
        df_usd_main[[f'Дата {usd_data["currency_pair"]}', f'Курс {usd_data["currency_pair"]}',
                     f'Время {usd_data["currency_pair"]}']],
        df_jpy_main[[f'Дата {jpy_data["currency_pair"]}', f'Курс {jpy_data["currency_pair"]}',
                     f'Время {jpy_data["currency_pair"]}']],
        left_on=f'Дата {usd_data["currency_pair"]}',
        right_on=f'Дата {jpy_data["currency_pair"]}',
        how='inner'
    )

    result_df['Результат'] = result_df[f'Курс {usd_data["currency_pair"]}'] / result_df[
        f'Курс {jpy_data["currency_pair"]}']

    result_df = result_df[[
        f'Дата {usd_data["currency_pair"]}', f'Курс {usd_data["currency_pair"]}', f'Время {usd_data["currency_pair"]}',
        f'Дата {jpy_data["currency_pair"]}', f'Курс {jpy_data["currency_pair"]}', f'Время {jpy_data["currency_pair"]}',
        'Результат'
    ]]
    with pd.ExcelWriter(f'currency_rates_{jpy_data["from_date"]} - {jpy_data["to_date"]}.xlsx',
                        engine='xlsxwriter') as writer:
        result_df.to_excel(writer, sheet_name='Курсы валют', index=False)
        columns = [i for i in result_df.columns]
        workbook = writer.book
        worksheet = writer.sheets['Курсы валют']

        last_row = len(result_df) + 1
        worksheet.write(last_row, 0, 'Сумма:')
        worksheet.write_formula(last_row, 1, f'=SUM(B2:B{last_row})')
        worksheet.write_formula(last_row, 4, f'=SUM(E2:E{last_row})')
        worksheet.write_formula(last_row, 6, f'=SUM(G2:G{last_row})')

        money_format = workbook.add_format({'num_format': '#,##0.00₽'})
        money_format_index = [1, 4, 6]
        for idx, col in enumerate(columns):
            header_len = len(col) + 2
            if idx in money_format_index:
                worksheet.set_column(idx, idx, header_len, money_format)
            else:
                worksheet.set_column(idx, idx, header_len)
    return len(result_df)


def send_email(rows: int, from_date: str, to_date: str) -> None:
    from_email = os.getenv('EMAIL_FROM')
    to_email = os.getenv('EMAIL_TO')
    password = os.getenv('PASSWORD')
    smtp_server = "smtp.yandex.ru"

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = f"Данные курсов валют с MOEX за период с {from_date} по {to_date}"

    last_digit = rows % 10
    if rows % 100 in (11, 12, 13, 14):
        rows_text = f"{rows} строк"
    elif last_digit == 1:
        rows_text = f"{rows} строка"
    elif 2 <= last_digit <= 4:
        rows_text = f"{rows} строки"
    else:
        rows_text = f"{rows} строк"

    body = f"""Данные по курсам валют с MOEX за период с {from_date} по {to_date}.
В прикрепленном файле содержится {rows_text}."""

    msg.attach(MIMEText(body, 'plain'))

    filename = f"currency_rates_{from_date} - {to_date}.xlsx"
    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {filename}")
    msg.attach(part)
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL(smtp_server, 465, context=context) as server:
            server.login(from_email, password)
            server.sendmail(from_email, to_email, msg.as_string())
        print("Письмо успешно отправлено!")
    except Exception as e:
        print(f"Ошибка при отправке письма: {e}")


send_email(create_excel(fetch_currency_data(currency_pair_left, from_date_input, to_date_input),
                        fetch_currency_data(currency_pair_right, from_date_input, to_date_input)), from_date_input,
           to_date_input)
