import win32com.client
import json
import os

nanocad_app = win32com.client.Dispatch("nanoCADx64.Application.22.0")  # "Проверка запуска программы nanoCAD.Application"
if nanocad_app is not None:
    ncad_doc = nanocad_app.ActiveDocument   # "Проверка открытия файла nanoCAD.Application"
    if ncad_doc is not None:
        for one_layout_index in range(0, ncad_doc.Layouts.Count, 1):  # "Число листов"
            ncad_Layout = ncad_doc.Layouts.Item(one_layout_index)  # "Сами листы"
            if ncad_Layout.Name != "Model":  # "Лист не Модель"
                ncad_Block_for_Layout = ncad_Layout.Block  # "Получение всех блоков на листе"
                # Создание основной надписи по центру листа
                # Стиль текста будет текущим в нанокаде, так что предварительно если нужно меняйте там

                def insert_title_as_text(text_to_inserting):  # "Пишем функцию вставки текста"
                    size_of_layout = ncad_Layout.GetPaperSize()  # "Получение размеров листа (ширина и высота в мм)"
                    center_text = str(size_of_layout[0]/2) + "," + str(size_of_layout[1]-25) + ",0"  # "Получение координат вставки текста (по центру по х по у на 25 мм ниже)"
                    title_text_inst = ncad_Block_for_Layout.AddText(text_to_inserting, center_text, 5)  # "Вставка текста по координатам, высота текста 5"
                    pass

                insert_title_as_text("Схема участка дороги")  # "Собсна сама вставка текста с использованием нашей функции"

                # Функция заполнения шапки

                def modify_dyn_block(block_instance):  # "Пишем функцию заполнения шапки"
                    to_change = {  # "Словарь изменений"
                        "РАЗРАБОТАЛ_ФАМИЛИЯ": "Гребенюк",
                        "РУКОВОДИТЕЛЬ_ФАМИЛИЯ": "Рудской",
                        "ДАТА": "03.22",
                        "1": "03.22",
                        "2": "03.22",
                        "3": "03.22",
                        "4": "03.22",
                    }
                    for one_attr in block_instance.GetAttributes():  # "Для каждого атрибута нашего блока"
                        if one_attr.TextString in to_change.keys():  # "Если атрибут есть в нашем словарике"
                            one_attr.TextString = to_change[one_attr.TextString]  # "То меняем значение этого атрибута"
                    pass

                for object_at_block_index in range(0,ncad_Block_for_Layout.Count, 1):  # "Индексы всех объектов в блоке"
                    object_at_block = ncad_Block_for_Layout.Item(object_at_block_index)  # "Получение объекта в блоке по индексу"
                    if object_at_block.ObjectName == "AcDbBlockReference":  # "Ищем динамический блок"
                        if object_at_block.EffectiveName == "штамп_Макорус":  # "Ищем блок с нужным именем"
                            modify_dyn_block(object_at_block)  # "Юзаем нашу функцию

                # Вставка в виде таблицы
                # ОТКРЫВАЕМ ЭКСЕЛЬ просто программу
                # Нужно сделать лист, где 1 столбик - наименование, 2 - количество, вроде бы заголовки нам не нужны, поэтому их не будет в примере ниже
                example = [[1, 2], [3, 4]]  # Допустим это наш список

                current_dir_path = os.getcwd()  # Создание временного эксель файла
                sample_excel_path = os.path.join(current_dir_path, 'test_excel.xlsx')

                def ms_excel_work():
                    excel_app = win32com.client.Dispatch("Excel.Application")
                    excel_app.DisplayAlerts = False
                    excel_app.Visible = True
                    working_file = excel_app.Workbooks.Add()
                    active_sheet = working_file.ActiveSheet
                    for one_rows_index in range(0, len(example), 1):
                        one_row = example[one_rows_index]
                        for prop_value_index in range(0, len(one_row), 1):
                            cell_working = active_sheet.Cells(one_rows_index + 2, prop_value_index + 1)
                            cell_working.Value = one_row[prop_value_index]
                    working_file.SaveAs(sample_excel_path)
                    # working_file.Close()
                    # excel_app.Quit()
                    pass


                ms_excel_work()

    else:
        print("Doc is not running")
else:
    print("App is not running")
