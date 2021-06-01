# Продавцы присылают каждый день EXCEL отчет о продажах из 1С
# Эта программа формирует файл CSV для загрузки в Интернет-магазин

import pandas as pd

import xlrd

from os import path

from django.http import HttpResponse


def separate_sku_from_size(sku):
    # Артикул вида 220GSBC2303-140*72*63  разбирает и на выходе отдает
    # our_sku1 = 220GSBC2303
    # our_size1 = 140
    # case = 'NORMAL SIZE'

    my_sku = sku.strip()  # удаляем пробелы сначала и в конце

    if my_sku.count('-') > 0:  # если тире есть, то выполняем, а если нет тире то ариткул и есть артикул

        splitted_word = my_sku.split('-')  # хорошая штука, делит строку на две части, разделитель тире

        our_sku1 = splitted_word[0]  # левая часть артикула (сам артикул)
        our_size1 = splitted_word[1]  # правая часть артикула (размер)
        our_size1_lowercase = our_size1.lower()

        if our_size1.count('*') == 2:  # если в размере две звездочки
            t = our_size1.find('*')  # находим номер символа первой звездочки
            our_size1 = our_size1[0:t]  # сокращаем размер - отбрасываем первую звездочку и все после нее
            case = 'NORMAL SIZE'
        elif our_size1.count('*') == 1:  # если одна звездочка в размере
            t = our_size1.find('*')  # то может быть два варианта - либо это тоддлеры, либо спаренные размеры

            temp_size_divided = our_size1.split('*')  # делим размер на две части , разделитель звездочка
            left_part_of_size = int(temp_size_divided[0])
            right_part_of_size = int(temp_size_divided[1])

            if left_part_of_size > right_part_of_size:
                our_size1 = our_size1[0:t]  # считаем так: если размер типа 92*52 то оставим только 92
                case = 'TODDLER SIZE'
            else:
                case = 'DOUBLE SIZE'
            # а если 14*16 то оставляем без изменений
        elif our_size1_lowercase.find('size') != -1: # Если No size или One size
            our_size1 = ''
            case = 'NO SIZE'
        else:
            case = 'ONE DIGIT SIZE'
    else:
        our_sku1 = my_sku
        our_size1 = ''
        case = 'NO SIZE'

    return our_sku1, our_size1, case






def handle_sales_report(excel_file, csv_file):

    def find_first_row_with_sku():
        # Найдем первую рабочую строку с артикулом
        # Критерий - чтобы артикул начинался с первых трех цифр
        x = 0
        for x in range(0, sheet.nrows):
            work_cell = str(sheet.cell_value(x, 0))
            # work_cell2 = int(sheet.cell_value(x, 1))
            try:
                if work_cell[0:3].isdigit():
                    break
            except:
                pass
        return x

    def already_written(sku):
        n = new_data[new_data['sku'] == sku].index.values

        if n.size == 0:
            return 0
        else:
            return n[0]

    def have_variants(m1):  # проверим в CSV файле есть ли варианты у этого артикула
        try:  # try - except сделан на случай если строка m1+1 не существует и там конец файла
            if my_csv_file.at[m1 + 1, 'fieldType'] == 'Variant':
                return True
            else:
                return False
        except:
            return False

    # global excel_file
    # global csv_file


    # result_file = 'READY_TO_WIX.csv'


    # text='ЗАПУСТИТЬ ОБРАБОТКУ ФАЙЛА'
    # text='Результат будет записан в .csv файл начинающимся на "READY TO WIX" '



    file = excel_file
    file_csv = csv_file

    # Если файл XLSX выглядит как МЕГА 1С 2020-10-11, то возьмем часть с датой и используем ее в имени результирующего файла
    # pos1 = excel_file.rfind('/')
    pos2 = excel_file.name.find('.')
    # result_file='READY TO WIX ' + excel_file.name[:pos2] + '.csv'

    # print('Result file name', result_file)

    wb = xlrd.open_workbook(file_contents=excel_file.read())

    sheet = wb.sheet_by_index(0)  # берем лист из экселевского файла

    # прочитаем файл csv и загрузим его в Датафрейм
    # столбец visible должен быть стринг, а не булеан, ато Wix может ругаться
    my_csv_file = pd.read_csv(file_csv, delimiter=',', index_col=False, encoding='utf-8', dtype={'visible': str, 'sku': str})


    # создадим ка новый пустой DataFrame
    col = my_csv_file.columns.values  # первая строка заголовков
    new_data = pd.DataFrame(columns=col)  # создаем пустой датафрейм с нашими заголовками

    main_log = {}
    main_log['Итого'] = {}
    inside_log={}

    count_return = 0
    count_cantfind = 0
    count_products = 0
    count_old_collection = 0
    count_morethanonce = 0
    count_lessthanzero = 0
    sku_cantfind_list = []
    sku_lessthanzero_list = []

    first_work_row = find_first_row_with_sku()     # находим первую строку с артикулом, начинаем работу с нее
    print('first work row', first_work_row)
    print('sheet.nrows', sheet.nrows)
    i = 1  # счетчик строк в записываемом итоговом датафрейме
    a = first_work_row


    for a in range(first_work_row, sheet.nrows):  # главный цикл чтобы пробегать по всем строкам.

        work_cell = str(sheet.cell_value(a, 0))  # читаем артикул
        quantity_cell = sheet.cell_value(a, 1)  # читаем количество проданного
        # print('------------------')
        # print(work_cell)
        # print(quantity_cell, type(quantity_cell))

        sku_long = work_cell.strip()  # это артикул целиком вместе с размером

        if not sku_long[:3].isdigit() or quantity_cell == 0 or quantity_cell=='':
            continue        # чтобы пропустить строку где нет артикула или количество не цифра или ноль

        sku = sku_long # в дальнейшем будем использовать чисто sku , по умолчанию пусть будет длинное sku

        quantity_cell=int(quantity_cell)

        sku_short, our_size, our_case = separate_sku_from_size(work_cell)
        main_log[sku_short]={}

        print('-----------------------')
        print('Артикул:', sku_short, ' Размер:', our_size)          # это артикул очищенный от размера

        print('Тип размера:', our_case)

        print('Количество проданного товара:', quantity_cell)

        if quantity_cell < 0:
            print('Возврат!')
            count_return += 1

        main_log[sku_short].update({'Размер': our_size})
        main_log[sku_short].update({'Количество продано': quantity_cell})

        # if quantity_cell == 0:
        #     print('Нет продажи, записываем без изменений!')
        #     count_nochange += 1

        m = my_csv_file[my_csv_file['sku'] == sku_short].index.values  # найдем индекс строки с нашим артикулом

        if m.size == 0:
            m = my_csv_file[my_csv_file['sku'] == sku_long].index.values  # попробуем ка второй вариант :
            #    ищем ПОЛНЫЙ АРТИКУЛ включая размер. Старые коллекции так у нас заведены
            if m.size != 0 :
                sku = sku_long
                count_old_collection += 1  # значит есть размер и нет вариантов и артикул надо использовать sku_long
        else :
            sku = sku_short

        if m.size == 0:     #если и полный артикул не нашли, значит точно нет такого в CSV файле
            print('Не нашли такой SKU в интернет-магазине!')
            main_log[sku_short].update({'Статус': 'Не нашли такой SKU в интернет-магазине!'})
            count_cantfind += 1
            sku_cantfind_list.append(sku_long)
            continue        # пропускаем

        # ======================= Все, такой артикул точно есть в CSV-файле =============================
        #далее используем только sku (никаких sku_short или sku_long)

        m1 = m[0]  # номер строки
        print('Номер строки с которой будем работать:', m1)

        if our_case != 'NO SIZE' and have_variants(m1) : #есть размер и есть варианты
        # if our_case != 'NO SIZE' and sku == sku_short:  # есть размер и есть варианты
            n = already_written(sku)
            if n == 0:  # ноль - значит не записывали еще
                new_data.loc[i] = my_csv_file.loc[m1]  # записали первую строку Product
                count_products += 1

                m1 = m1 + 1  # перешли на следующую строку в CSV файле интернет-магазина
                i = i + 1  # перешли на следующую строку в новом CSV файле
                while my_csv_file.at[m1, 'fieldType'] == 'Variant':  # пока ниже идут поля "Variant"

                    new_data.loc[i] = my_csv_file.loc[m1]  # записываем строку

                    if my_csv_file.at[
                        m1, 'productOptionDescription1'] == our_size:  # если размер равен нашему размеру

                        z_str = new_data.at[i, 'inventory']
                        z_int = int(z_str)

                        print('Количество товара:', z_int,' ---> ',z_int-quantity_cell)
                        main_log[sku_short].update({'Было': z_int})
                        main_log[sku_short].update({'Стало': z_int-quantity_cell})

                        z_int = z_int - quantity_cell
                        z_str = str(z_int)
                        if z_int < 0 :
                            count_lessthanzero +=1
                            sku_lessthanzero_list.append(sku_long)

                        new_data.at[i, 'inventory'] = z_str  # нам надо уменьшить количество запаса товара на quantity

                    m1 = m1 + 1
                    i = i + 1
            else:   # если n - не ноль, значит записывали уже
                print('Повторный случай')
                count_morethanonce +=1
                n=n+1
                try:
                    while new_data.at[n, 'fieldType'] == 'Variant' :  # пока ниже идут поля "Variant"

                        if new_data.at[n, 'productOptionDescription1'] == our_size:  # если размер равен нашему размеру

                            z_str = new_data.at[n, 'inventory']
                            z_int = int(z_str)

                            print('Старое количество товара:', z_int)
                            main_log[sku_short].update({'Было': z_int})

                            z_int = z_int - quantity_cell
                            z_str = str(z_int)
                            if z_int < 0:
                                count_lessthanzero += 1
                                sku_lessthanzero_list.append(sku_long)
                            new_data.at[n, 'inventory'] = z_str  # нам надо уменьшить количество запаса товара на quantity

                            print('Новое количество: ', new_data.at[n, 'inventory'])
                            main_log[sku_short].update({'Стало': z_int})

                        n = n + 1
                except : # когда конец new_data доcтигнут, чтобы не вылетал с ошибкой
                    print('конец new_data!')

        else:   # есть размер и нет вариантов  ЛИБО нет размера - вобщем только строка Product и все
            n = already_written(sku)
            if n == 0: # ноль - значит не записывали еще
                new_data.loc[i] = my_csv_file.loc[m1]  # записали первую строку Product
                z_str = new_data.at[i, 'inventory']
                z_int = int(z_str)
                print('Было количество:', z_int, ' стало: ', z_int - quantity_cell)
                main_log[sku_short].update({'Было': z_int})
                main_log[sku_short].update({'Стало': z_int - quantity_cell})

                z_int = z_int - quantity_cell
                z_str = str(z_int)
                if z_int < 0:
                    count_lessthanzero += 1
                    sku_lessthanzero_list.append(sku_long)
                new_data.at[i, 'inventory'] = z_str     # записали новое количество в new_data

                i += 1  # перешли на след. строку в new_data Dataframe
                m1 += 1 # перешли на след. строку в csv файле  А НУЖНО ЛИ ЭТО ЗДЕСЬ?
                count_products += 1

            else:   # уже есть запись в new_data по номеру строки
                print('Повторный случай')
                count_morethanonce += 1
                z_str = new_data.at[n, 'inventory']
                z_int = int(z_str)
                print('Старое количество товара:', z_int)
                main_log[sku_short].update({'Было': z_int})

                z_int = z_int - quantity_cell
                print('Новое количество: ', z_int)
                main_log[sku_short].update({'Стало': z_int})
                z_str = str(z_int)
                if z_int < 0:
                    count_lessthanzero += 1
                    sku_lessthanzero_list.append(sku_long)
                new_data.at[n, 'inventory'] = z_str  # записали новое количество в new_data


    print('***********************************')
    print('Products:',count_products)
    print('Плюс к ним повторных случаев:', count_morethanonce)
    print('Возвратов:',count_return)
    print('Не найдено:', count_cantfind, sku_cantfind_list)
    print('Из аутлета (из найденных в ИМ):', count_old_collection)
    print('Кол-во меньше нуля (нужно исправить):',count_lessthanzero, sku_lessthanzero_list)

    main_log['Итого'].update({'Количество товаров': count_products,
                              'Возвратов': count_return,
                              'Не найдено': count_cantfind,
                              'Из аутлета': count_old_collection,
                              'Кол-во меньше нуля (нужно исправить)': count_lessthanzero
                              })
    # main_log['Итого'].update({'Возвратов': count_return})



    return new_data, main_log

