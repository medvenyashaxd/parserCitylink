import time
import requests
import xlsxwriter

from bs4 import BeautifulSoup


def search_data():

    number_page = 1  # Счетчик страниц
    session = requests.Session()
    header = {
        'authority': 'www.citilink.ru',
        'method': 'GET',
        'path': f'/catalog/blendery/?p={number_page}',
        'scheme': 'https',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'ru,en-US;q=0.9,en;q=0.8,ru-RU;q=0.7',
        'cache-control': 'max-age=0',
        'cookie': 'old_design=0; is_show_welcome_mechanics=1; _tuid=6d7d850cfa08c723c73a7ebb9c4e8afa31ff898b; _space=msk_cl%3A; ab_test_main_10_10_80=3; _slid=6388bc28ad014f1fbe0ab13c; _gcl_au=1.1.853433464.1669905450; _gid=GA1.2.486207756.1669905451; tmr_lvid=1bb2dc45185720f2b090a09b85e721b3; tmr_lvidTS=1669905450850; _ym_uid=16699054511040816842; _ym_d=1669905451; _userGUID=0:lb56l8df:4s4UAmi0tHKmbcdiDElQrd~7TlRcyo6n; advcake_session_id=6bca3d0a-737b-f39b-b0b5-a55320e7b87c; advcake_track_url=https%3A%2F%2Fwww.citilink.ru%2F; advcake_utm_partner=; advcake_utm_webmaster=; advcake_click_id=; _ym_isad=1; _tt_enable_cookie=1; _ttp=da41bcec-87f0-4885-81b1-b8677d8c92c7; __exponea_etc__=fedd8d35-6948-4f2e-bb94-778360b3b8c7; clientId=660220997.1669905451; _slid_server=6388bc28ad014f1fbe0ab13c; _dy_ses_load_seq=44395%3A1669912491751; _dy_c_exps=; _dy_soct=1017570.1030352.1669912491*1033770.1068198.1669912491*1036008.1075335.1669912491*1008131.1012968.1669912491*1015299.1026208.1669912491; ab_test_productcard_5_5_90=3; advcake_track_id=1a2b3415-8973-7322-c980-b31727e9ecbe; _pcl=eW5jicFlc6lHyQ==; _slsession=AA75D70A-5C92-4866-877C-D680B7665C1B; _slfreq=6347f312d9062ed0380b52dc%3A6347f38c9a3f3b9e90027775%3A1669979576; _dvs=0:lb6afo8q:g~wYr74HCaG3VkMl63tr1OmGCA6fih4b; _ym_visorc=w; AMP_TOKEN=%24NOT_FOUND; __exponea_time2__=-0.9911403656005859; dSesn=eb19883e-ba1d-d20c-2a49-3e92cb584a25; digi_uc=W1sidiIsIjE0NTc3ODEiLDE2Njk5NzQ5MzMxNDldLFsidiIsIjE4NDQwNDkiLDE2Njk5NzQ0NDk5OTVdLFsidiIsIjE0NzY4MjIiLDE2Njk5NzM4ODc5NjVdLFsidiIsIjEwNTEyMTIiLDE2Njk5MjY4MzE1MDhdLFsidiIsIjE0NzY4MjgiLDE2Njk5MjUyNjkyMzNdLFsidiIsIjE2ODg3ODEiLDE2Njk5MTgwOTY4NzVdLFsidiIsIjE4NDg0NDQiLDE2Njk5MTI1NTE4MjddLFsiY3YiLCIxNDU3NzgxIiwxNjY5OTc0ODE4NDQ4XSxbImN2IiwiNzIyODgwIiwxNjY5OTc0MzQzMjkxXSxbImN2IiwiMTQxOTQ2NSIsMTY2OTk3MzM0NDI0OV0sWyJjdiIsIjE2MDMyMzAiLDE2Njk5NzMzMDQwNTRdLFsiY3YiLCIxNDI4NzQwIiwxNjY5OTcyNTI3MzY0XSxbImN2IiwiMTQxOTQwMCIsMTY2OTk3MjUyMzU2OF0sWyJjdiIsIjE0Mjg3MjQiLDE2Njk5NzI1MjExMTldLFsiY3YiLCIxNjA3MzY2IiwxNjY5OTcyNTE4OTU1XSxbImN2IiwiODYwNDQxIiwxNjY5OTcyNDU5NTM3XSxbImN2IiwiMTg3MjQyNiIsMTY2OTkxODE1MDk2MV1d; rr_rcs=v%3A1457781%3A1669974933233%3Bv%3A1844049%3A1669974450219%3Bv%3A1476822%3A1669973888152%3Bv%3A1051212%3A1669926831877%3Bv%3A1476828%3A1669925269607%3Bv%3A1848444%3A1669912552231; ab_test=90x10v4%3A1%7Creindexer%3A2%7Cdynamic_yield%3A3%7Cwelcome_mechanics%3A4%7Cdummy%3A10%7Cpage_listing%3A3; ab_test_analytics=90x10v4%3A1%7Creindexer%3A2%7Cdynamic_yield%3A3%7Cwelcome_mechanics%3A4%7Cdummy%3A10%7Cpage_listing%3A3; mindboxDeviceUUID=397ccf28-9105-44a4-b844-ac6135deaeb0; directCrm-session=%7B%22deviceGuid%22%3A%22397ccf28-9105-44a4-b844-ac6135deaeb0%22%7D; tmr_detect=1%7C1669975704489; _ga=GA1.2.660220997.1669905451; _ga_DDRSRL2E1B=GS1.1.1669972376.6.1.1669975737.6.0.0',
        'referer': 'https://www.citilink.ru/catalog/melkaya-bytovaya-tehnika/',
        'sec-ch-ua': '"Google Chrome";v="107", "Chromium";v="107", "Not=A?Brand";v="24"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36',
    }

    while number_page != 10 + 1:  # Кол-во страниц для парсинга
        url = f'https://www.citilink.ru/catalog/blendery/?p={number_page}'  # Ссылка на ресурс
        response_catalog = session.get(url, headers=header)  # Получаем html каталога
        print(response_catalog)  # Код ответа от сервера

        soup_catalog = BeautifulSoup(response_catalog.text, 'lxml')  # Обрабатываем парсером
        products = soup_catalog.findAll('a', class_='ProductCardHorizontal__title Link js--Link Link_type_default')  # Находим товары в каталоге

        for product in products:  # В цикле разбираем каждый товар
            href = 'https://www.citilink.ru/' + product.get('href')  # Получаем ссылку на товар
            response_card = session.get(href, headers=header)  # Получаем html товара
            soup_card = BeautifulSoup(response_card.text, 'lxml')  # Обрабатываем товар парсером
            properties = href + 'properties/'  # Получаем ссылку для перехода в характеристики товара
            response_properties = session.get(properties, headers=header)  # Получаем HTML характеристик товара
            soup_properties = BeautifulSoup(response_properties.text, 'lxml')  # Обрабатываем свойства парсером

            try:  # Пробуем получить данные
                product_id = soup_card.find('div', class_='ProductHeader__product-id').text.split(':')[1].strip()  # Получаем код товара
                brand = soup_properties.find('div', class_='Specifications__column Specifications__column_value').text.strip()
                    # \
                    # .findNext('div', class_='Specifications__column Specifications__column_value')\
                    # .findNext('div', class_='Specifications__column Specifications__column_value').text.strip()  # Получаем наименование бренда

                model = soup_properties.find('div', class_='Specifications__column Specifications__column_value').findNext('div', class_='Specifications__column Specifications__column_value').text.strip()
                    # .findNext('div', class_='Specifications__column Specifications__column_value')\
                    # .findNext('div', class_='Specifications__column Specifications__column_value').text.strip()  # Получаем наименование модели

                color = soup_properties.find('h1', class_='Heading Heading_level_1 ProductHeader__title').text.split(',')[2].strip()  # Получаем цвет модели

                photography = []
                photos = soup_card.findAll('img', class_='__image-upper PreviewList__image Image')  # Получаем фотографии
                for p in photos:
                    photography.append(p.get('src').replace('[', '').replace(']', '').replace("'", ''))


                # characteristics_name = soup_card.findAll('div', class_='Specifications__column Specifications__column_name')
                # characteristics_value = soup_card.findAll('div', class_='Specifications__column Specifications__column_value')
                # for name in characteristics_name:
                #     specification.append(name.text.replace('/n', '').replace(',', '').strip())
                #     for value in characteristics_value:
                #         specification.append(value.text.replace('/n', '').replace(',', '').strip())

                specification = []
                characteristics_value = soup_properties.findAll('div', class_='Specifications__column Specifications__column_value')  # Получаем характеристики
                for value in characteristics_value:
                    specification.append(value.text.replace('/n', '').replace(',', '').replace('[', '').replace(']', '').replace("'", '').strip())

                print(product_id, brand, model, color, photography, specification)
                yield product_id, brand, model, color, photography, specification  # Передаем данные для записи в Exel таблицу

            except Exception as error:  # Обработка ошибок
                print(error)
                time.sleep(3)

        number_page += 1  # После цикла переходим на следующую страницу
        print(number_page)


def start_pars():

    book = xlsxwriter.Workbook("C:\\Users\\xmedv\\Desktop\\Blenders.xlsx")  # Создаем файл Exel
    page = book.add_worksheet('Microwaves')  # Создаем лист внутри файла

    row = 0  # Счетчик строк
    column = 0  # Счетчик столбцов

    # Создание столбцов
    page.set_column('A:A', 15)
    page.set_column('B:B', 20)
    page.set_column('C:C', 9)
    page.set_column('D:D', 30)
    page.set_column('E:E', 14)
    page.set_column('F:F', 10)
    page.set_column('G:G', 30)
    page.set_column('H:H', 10)
    page.set_column('I:I', 20)


    # Создание описания столбцов
    page.write(0, 0, 'ID')
    page.write(0, 0+1, 'Раздел')
    page.write(0, 0+2, 'Бред')
    page.write(0, 0+3, 'Модель')
    page.write(0, 0+4, 'Фото')
    page.write(0, 0+5, 'Описание')
    page.write(0, 0+6, 'Скл. модель')
    page.write(0, 0+7, 'Продукт')
    page.write(0, 0+8, 'Артикул')
    row += 1

    for data in search_data():  # Производим запись данных в файл
        page.write(row, column, data[0])
        page.write(row, column+1, 'МБТ - Блендеры')
        page.write(row, column+2, data[1])
        page.write(row, column+3, data[2] + ' ' + data[3])
        page.write(row, column+4, str(data[4]).replace('[', '').replace(']', '').replace("'", ''))
        page.write(row, column+5, str(data[5]).replace('[', '').replace(']', '').replace("'", ''))
        page.write(row, column+7, 'Блендер')
        page.write(row, column+8, data[2] + ' ' + data[3])
        row += 1

    book.close()


if __name__ == '__main__':
    start_pars()

