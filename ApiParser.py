# Импортируем библиотеки
import requests
import openpyxl
# Время работы программы
import time
time_string_start = time.strftime("%m/%d/%Y, %H:%M:%S")
print(time_string_start)
t0 = time.time()
book = openpyxl.Workbook()
sheet = book.active
row = 2
sheet[f'A1'] = "Название"
sheet[f'B1'] = "Код АТИ"
sheet[f'C1'] = "Город"
sheet[f'D1'] = "ИНН"
sheet[f'E1'] = "ОГРН"
sheet[f'F1'] = "Контакт"
sheet[f'G1'] = "Телефон"
sheet[f'H1'] = "Моб. Телефон"
sheet[f'I1'] = "Почта"
sheet[f'J1'] = "Балл АТИ"
sheet[f'K1'] = "Рекомендации"
start = int(input('Стартовая позиция диапазона парсинга: '))
end = int(input('Конечная позиция диапазона парсинга: '))
# Цикл перебора кодов АТИ
for num in range(start, end + 1):
    ati_id = num
    # Отправляем запрос на полчение данных фирмы по ее АТИ коду
    response = requests.get(f'https://api.ati.su/v1.0/firms/summary/{ati_id}',
                            headers={'Authorization': 'Bearer d4a87f8c7c5344da830f8d0f1528abda'})
    response1 = requests.get(f'https://api.ati.su/v1.0/firms/{ati_id}/contacts/summary',
                             headers={'Authorization': 'Bearer d4a87f8c7c5344da830f8d0f1528abda'})

    if response.status_code == 200:
        if response1.status_code == 200:
            response1 = response1.json()
            if len(response1) > 0:
                response = response.json()
                sheet[f'A{row}'] = response["full_name"]
                sheet[f'B{row}'] = response["ati_id"]
                sheet[f'C{row}'] = response["location"]["city_name"]
                sheet[f'D{row}'] = response["inn"]
                sheet[f'E{row}'] = response["ogrn"]
                sheet[f'J{row}'] = response["score"]
                sheet[f'K{row}'] = response["recommendations_count"]
                sheet[f'F{row}'] = response1[0]["name"]
                sheet[f'G{row}'] = response1[0]["phone"]
                sheet[f'H{row}'] = response1[0]["mobile_phone"]
                sheet[f'I{row}'] = response1[0]["email"]

                # for contact in range(len(response1)):
                #     sheet[f'F{row}'] = response1[contact]["name"]
                #     sheet[f'G{row}'] = response1[contact]["phone"]
                #     sheet[f'H{row}'] = response1[contact]["mobile_phone"]
                #     sheet[f'I{row}'] = response1[contact]["email"]
                row += 1
        book.save(f'Base {start} - {end}.xlsx')
print('Завершено')
time_string_end = time.strftime("%m/%d/%Y, %H:%M:%S")
print(time_string_end)
t1 = time.time()
print("Затраченное время: ", t1 - t0)
book.close()
