import openpyxl
import os
import shutil
# Скрипт для формирования опросных листов датчиков КИП
# Версия программы:
Version = '1.2'

def saving(file_name):
    if not os.path.exists(UserPath + '/' + file_name):
        sh.save(UserPath + '/' + file_name)
        print('Опросный лист на датчик ' + Designator + Position + ' создан успешно!')
    else:
        print(
            'Опросный лист на датчик ' + Designator + Position + ' существует. Заменить опросный лист с повторяющейся позицией? (Да/Нет)')
        podtverjdenie = input()
        if podtverjdenie == 'Да' or podtverjdenie == 'да' or podtverjdenie == 'ДА' or podtverjdenie == 'дА':
            sh.save(UserPath + '/' + file_name)
            print('Опросный лист на датчик ' + Designator + Position + ' заменен!')
        else:
            sh.save(UserPath + '/' + 'Дублированый ' + file_name)
            print('Создан опросный лист на датчик с повторяющейся позицией ' + Designator + Position)

    sh.close()

def exd_type(type):
    if Exd == 'Ex' or Exd == 'да' or Exd == 'Exd':
        list1[type].value = 'Взрывозащищенное (Exd)'  # Вид исполнения
    else:
        list1[type].value = 'Общепромышленное'
def exd_entry(entry):
    if Exd == 'Ex' or Exd == 'да' or Exd != '-':
        list1[entry].value = 'Да'  # Взрывозащищенность кабельного ввода
    else:
        list1[entry].value = 'Нет'

print('Версия программы: ', Version, ' от 12.10.2022')
print('При копировании пути к исходному файлу, следите, чтобы не было ковычек!')
print('Введите путь к исходному Перечню КИП:')
book = input()
print('Куда сохранить созданные опросные листы? Введите путь:')
Path = input()
print('Введите имя папки для опросных листов:')
dirName = input()
UserPath = Path + '/' + dirName
if not os.path.exists(UserPath):
       os.mkdir(UserPath)
else:
    print('Папка с таким именем уже существует! Введите другое имя папки:')
    dirName = input()
    UserPath = Path + '/' + dirName
    os.mkdir(UserPath)

wb = openpyxl.load_workbook(book)
sheet = wb.active
massive = [0] * sheet.max_column
counter = 0

for stroka in range(3, sheet.max_row + 1):
    i = 0
    massive = [0] * sheet.max_column
    for row in range(0, sheet.max_column):
        massive[i] = sheet[stroka][row].value
        i = i + 1

    Otdelenie = massive[0]
    Type_device = massive[2]
    Location = massive[3]
    Isolation = massive[4]
    Environment = massive[5]
    Designator = str(massive[6])
    Position = str(massive[7])
    Range = massive[8]
    Scale = massive[9]
    Unit = massive[10]
    Low = massive[11]
    PreLow = massive[12]
    PreHigh = massive[13]
    High = massive[14]
    Pressure = massive[15]
    Temperature = massive[16]
    ExplTemp = massive[17]
    Density = massive[18]
    Viscosity = massive[19]
    Special_properties = massive[20]
    OutSignal = massive[21]
    Error = massive[26]
    Exd = massive[27]
    ExdEnv = massive[28]
    Connection = massive[29]
    ProcessMaterial = massive[30]
    Note = massive[33]
    IP = massive[34]
    UHL = massive[35]
    L = massive[36]
    ValveBlock = massive[37]
    Cooler = massive[38]
    Separator = massive[39]
    Cartridge = massive[40]
    LED = massive[41]
    Screen = massive[42]
    Supply = massive[43]
    CableEntry = massive[44]
    LegsPosition = massive[45]
    LegsQuantity = massive[46]
    VesselPosition = massive[47]
    VesselMaterial = massive[48]

    if Unit != None:
        if Designator == 'TT' or Designator == 'ТТ' or Designator == 'TТ' or Designator == 'ТT':
            sh = openpyxl.load_workbook('TT.xlsx')
            list1 = sh.active
            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Range  # Диапазон измерения
            list1['F22'].value = Error  # Абсолютная погрешность измерения
            list1['F23'].value = Pressure  # Давление среды
            list1['F24'].value = Temperature  # Температура в точке монтажа
            list1['F25'].value = ExplTemp  # Температура окружающего воздуха
            list1['F26'].value = Density  # Плотность
            list1['F27'].value = Viscosity  # Вязкость
            list1['F28'].value = Special_properties  # Особые свойства
            list1['F29'].value = OutSignal  # Выходной сигнал
            list1['F30'].value = Supply  # Напряжение питания
            # Механические характеристики
            list1['F32'].value = Connection  # Соединение с процессом
            list1['F33'].value = Cartridge  # Гильза
            list1['F34'].value = ProcessMaterial  # Материальное исполнение в контакте с рабочей средой
            list1['F35'].value = Isolation  # Толщина теплоизоляции в месте установки
            list1['F36'].value = L # Длина монтажной части
            list1['F37'].value = IP  # Степень пылевлагозащиты IP
            list1['F38'].value = UHL  # Климатическое исполнение
            list1['F39'].value = CableEntry  # Тип кабельного ввода
            # Взрывозащита

            exd_type('F41')  # Вид исполнения
            exd_entry('F43')   # Взрывозащищенность кабельного ввода

  #          if Exd == 'Ex':
  #              list1['F41'].value = 'Взрывозащищенное (Exd)'  # Вид исполнения
  #          else:
  #              list1['F41'].value = 'Общепромышленное'

            list1['F42'].value = ExdEnv # Вещество для определения степени взрывозащиты

  #          if Exd == 'Ex':
  #              list1['F43'].value = 'Да'  # Взрывозащищенность кабельного ввода
  #          else:
  #              list1['F43'].value = 'Нет'

            # Интерфейс
            list1['F45'].value = LED  # Наличие индикации
            list1['F46'].value = Screen  # Лок. интерфейс оператора для настройки датчика
            # Примечание
            list1['A48'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1

        if Designator == 'AT' or Designator == 'AТ' or Designator == 'АТ' or Designator == 'АТ' or Designator == 'AA' \
                or Designator == 'AА' or Designator == 'АA' or Designator == 'АА':
            sh = openpyxl.load_workbook('AT.xlsx')
            list1 = sh.active
            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Range  # Диапазон измерения
            list1['E21'].value = Unit
            list1['F22'].value = Error  # Абсолютная погрешность измерения
            list1['E22'].value = Unit  # Абсолютная погрешность измерения
            list1['F23'].value = ExplTemp  # Температура окружающего воздуха
            list1['F24'].value = OutSignal  # Выходной сигнал
            list1['F25'].value = Supply  # Напряжение питания
            # Механические характеристики
            list1['F27'].value = IP  # Степень пылевлагозащиты IP
            list1['F28'].value = UHL  # Климатическое исполнение
            list1['F29'].value = CableEntry  # Тип кабельного ввода
            # Взрывозащита
            exd_type('F31')  # Вид исполнения
            exd_entry('F33')  # Взрывозащищенность кабельного ввода
            list1['F32'].value = ExdEnv  # Вещество для определения степени взрывозащиты

            # Интерфейс
            list1['F35'].value = LED  # Наличие индикации
            list1['F36'].value = Screen  # Лок. интерфейс оператора для настройки датчика
            # Примечание
            list1['A38'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1

        if Designator == 'FT' or Designator == 'FТ':
            sh = openpyxl.load_workbook('FT.xlsx')
            list1 = sh.active

            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Range  # Диапазон измерения
            list1['E21'].value = Unit #Единица измерения
            list1['F22'].value = Error  # Абсолютная погрешность измерения
            list1['E22'].value = Unit  # Единица измерения
            list1['F23'].value = Pressure  # Давление среды
            list1['F24'].value = Temperature  # Температура в точке монтажа
            list1['F25'].value = ExplTemp  # Температура окружающего воздуха
            list1['F26'].value = Density  # Плотность
            list1['F27'].value = Viscosity  # Вязкость
            list1['F28'].value = Special_properties  # Особые свойства
            list1['F29'].value = OutSignal  # Выходной сигнал
            list1['F30'].value = Supply  # Напряжение питания
            # Механические характеристики
            list1['F32'].value = Connection  # Соединение с процессом
            list1['F33'].value = ProcessMaterial  # Материальное исполнение в контакте с рабочей средой
            list1['F34'].value = Isolation  # Толщина теплоизоляции в месте установки
            list1['F35'].value = IP  # Степень пылевлагозащиты IP
            list1['F36'].value = UHL  # Климатическое исполнение
            list1['F37'].value = CableEntry  # Тип кабельного ввода
            # Взрывозащита
            exd_type('F39')  # Вид исполнения
            exd_entry('F41')  # Взрывозащищенность кабельного ввода

            list1['F40'].value = ExdEnv  # Вещество для определения степени взрывозащиты

            # Интерфейс
            list1['F43'].value = LED  # Наличие индикации
            list1['F44'].value = Screen  # Лок. интерфейс оператора для настройки датчика
            # Примечание
            list1['A46'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1

        if Designator == 'LS':
            sh = openpyxl.load_workbook('LS.xlsx')
            list1 = sh.active
            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Range  # Диапазон измерения
            list1['F22'].value = Pressure  # Давление среды
            list1['F23'].value = Temperature  # Температура в точке монтажа
            list1['F24'].value = ExplTemp  # Температура окружающего воздуха
            list1['F25'].value = Density  # Плотность
            list1['F26'].value = Viscosity  # Вязкость
            list1['F27'].value = Special_properties  # Особые свойства
            list1['F28'].value = OutSignal  # Выходной сигнал
            list1['F29'].value = Supply  # Напряжение питания
            # Механические характеристики
            list1['F31'].value = Connection  # Соединение с процессом
            list1['F32'].value = ProcessMaterial  # Материальное исполнение в контакте с рабочей средой
            list1['F33'].value = Isolation  # Толщина теплоизоляции в месте установки
            list1['F34'].value = L  # Длина монтажной части
            list1['F35'].value = IP  # Степень пылевлагозащиты IP
            list1['F36'].value = UHL  # Климатическое исполнение
            list1['F37'].value = CableEntry  # Тип кабельного ввода
            # Взрывозащита
            exd_type('F39')  # Вид исполнения
            exd_entry('F41')  # Взрывозащищенность кабельного ввода

            list1['F40'].value = ExdEnv  # Вещество для определения степени взрывозащиты

            # Интерфейс
            list1['F43'].value = LED  # Наличие индикации
            list1['F44'].value = Screen  # Лок. интерфейс оператора для настройки датчика
            # Примечание
            list1['A46'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1

        if Designator == 'LT' or Designator == 'LТ' or Designator == 'LDT' or Designator == 'LDТ':
            sh = openpyxl.load_workbook('LT.xlsx')
            list1 = sh.active
            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Range  # Диапазон измерения
            list1['E21'].value = Unit  # Единица измерения
            list1['F22'].value = Error  # Абсолютная погрешность измерения
            list1['E22'].value = Unit  # Единица измерения
            list1['F23'].value = Pressure  # Давление среды
            list1['F24'].value = Temperature  # Температура в точке монтажа
            list1['F25'].value = ExplTemp  # Температура окружающего воздуха
            list1['F26'].value = Density  # Плотность
            list1['F27'].value = Viscosity  # Вязкость
            list1['F28'].value = Special_properties  # Особые свойства
            list1['F29'].value = OutSignal  # Выходной сигнал
            list1['F30'].value = Supply  # Напряжение питания
            # Механические характеристики
            list1['F32'].value = Connection  # Соединение с процессом
            list1['F33'].value = ProcessMaterial  # Материальное исполнение в контакте с рабочей средой
            list1['F34'].value = IP  # Степень пылевлагозащиты IP
            list1['F35'].value = UHL  # Климатическое исполнение
            list1['F36'].value = CableEntry  # Тип кабельного ввода
            # Взрывозащита
            exd_type('F38')  # Вид исполнения
            exd_entry('F40')  # Взрывозащищенность кабельного ввода

            list1['F39'].value = ExdEnv  # Вещество для определения степени взрывозащиты

            # Интерфейс
            list1['F42'].value = LED  # Наличие индикации
            list1['F43'].value = Screen  # Лок. интерфейс оператора для настройки датчика
            # Примечание
            list1['A45'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1

        PTs = ['PT', 'РT', 'PТ', 'РТ', 'PDT', 'PDТ', 'PDТ', 'РDТ']

        if Designator in PTs:

            sh = openpyxl.load_workbook('PT.xlsx')
            list1 = sh.active
            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Range  # Диапазон измерения
            list1['E21'].value = Unit #Единица измерения
            list1['F22'].value = Error  # Абсолютная погрешность измерения
            list1['E22'].value = Unit  # Единица измерения
            list1['F23'].value = Pressure  # Давление среды
            list1['F24'].value = Temperature  # Температура в точке монтажа
            list1['F25'].value = ExplTemp  # Температура окружающего воздуха
            list1['F26'].value = Density  # Плотность
            list1['F27'].value = Viscosity  # Вязкость
            list1['F28'].value = Special_properties  # Особые свойства
            list1['F29'].value = OutSignal  # Выходной сигнал
            list1['F30'].value = Supply  # Напряжение питания
            # Механические характеристики
            list1['F32'].value = Connection  # Соединение с процессом
            list1['F33'].value = ProcessMaterial  # Материальное исполнение в контакте с рабочей средой
            list1['F34'].value = IP  # Степень пылевлагозащиты IP
            list1['F35'].value = UHL  # Климатическое исполнение
            list1['F36'].value = CableEntry  # Тип кабельного ввода
            # Взрывозащита
            exd_type('F38')  # Вид исполнения
            exd_entry('F40')  # Взрывозащищенность кабельного ввода

            list1['F39'].value = ExdEnv  # Вещество для определения степени взрывозащиты

            # Принадлежности
            list1['F42'].value = ValveBlock  # Клапанный блок
            list1['F43'].value = Cooler  # Охладитель
            list1['F44'].value = Separator  # Разделитель сред

            # Интерфейс
            list1['F46'].value = LED  # Наличие индикации
            list1['F47'].value = Screen  # Лок. интерфейс оператора для настройки датчика
            # Примечание
            list1['A49'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1

        if Designator == 'FS':
            sh = openpyxl.load_workbook('FS.xlsx')
            list1 = sh.active
            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Pressure  # Давление среды
            list1['F22'].value = Temperature  # Температура в точке монтажа
            list1['F23'].value = ExplTemp  # Температура окружающего воздуха
            list1['F24'].value = Density  # Плотность
            list1['F25'].value = Viscosity  # Вязкость
            list1['F26'].value = Special_properties  # Особые свойства
            list1['F27'].value = OutSignal  # Выходной сигнал
            list1['F28'].value = Supply  # Напряжение питания
            # Механические характеристики
            list1['F30'].value = Connection  # Соединение с процессом
            list1['F31'].value = ProcessMaterial  # Материальное исполнение в контакте с рабочей средой
            list1['F32'].value = Isolation  # Толщина теплоизоляции в месте установки
            list1['F33'].value = L  # Длина монтажной части
            list1['F34'].value = IP  # Степень пылевлагозащиты IP
            list1['F35'].value = UHL  # Климатическое исполнение
            list1['F36'].value = CableEntry  # Тип кабельного ввода
            # Взрывозащита
            exd_type('F38')  # Вид исполнения
            exd_entry('F40')  # Взрывозащищенность кабельного ввода

            list1['F39'].value = ExdEnv  # Вещество для определения степени взрывозащиты

            # Интерфейс
            list1['F42'].value = LED  # Наличие индикации
            list1['F43'].value = Screen  # Лок. интерфейс оператора для настройки датчика
            # Примечание
            list1['A45'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1

        if Designator == 'PG':
            sh = openpyxl.load_workbook('PG.xlsx')
            list1 = sh.active
            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Range  # Диапазон измерения
            list1['E21'].value = Unit #Единица измерения
            list1['F22'].value = Scale  # Шкала
            list1['E22'].value = Unit  # Единица измерения
            list1['F23'].value = Scale  # Абсолютная погрешность измерения
            list1['F24'].value = Pressure  # Давление среды
            list1['F25'].value = Temperature  # Температура в точке монтажа
            list1['F26'].value = ExplTemp  # Температура окружающего воздуха
            list1['F27'].value = Density  # Плотность
            list1['F28'].value = Viscosity  # Вязкость
            list1['F29'].value = Special_properties  # Особые свойства
            list1['F30'].value = OutSignal  # Электроконтакты
            # Механические характеристики
            list1['F32'].value = Connection  # Соединение с процессом
            list1['F33'].value = ProcessMaterial  # Материальное исполнение в контакте с рабочей средой
            list1['F34'].value = IP  # Степень пылевлагозащиты IP
            list1['F35'].value = UHL  # Климатическое исполнение
            # Взрывозащита
            exd_type('F37')  # Вид исполнения

            list1['F38'].value = ExdEnv  # Вещество для определения степени взрывозащиты

            # Принадлежности
            list1['F40'].value = ValveBlock  # Клапанный блок
            list1['F41'].value = Cooler  # Охладитель
            list1['F42'].value = Separator  # Разделитель сред

            # Примечание
            list1['A44'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1

        if Designator == 'WT' or Designator == 'WТ':
            sh = openpyxl.load_workbook('WT.xlsx')
            list1 = sh.active
            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Range  # Диапазон измерения
            list1['F22'].value = Error  # Абсолютная погрешность измерения
            list1['F23'].value = LegsPosition  # Расположение
            list1['F24'].value = LegsQuantity  # Количество опор
            list1['F25'].value = VesselPosition  # Положение емкости
            list1['F26'].value = VesselMaterial  # Материал емкости
            list1['F27'].value = ExplTemp  # Температура окружающего воздуха
            list1['F28'].value = OutSignal  # Выходной сигнал
            list1['F29'].value = Supply  # Напряжение питания
            # Механические характеристики
            list1['F31'].value = IP  # Степень пылевлагозащиты IP
            list1['F32'].value = UHL  # Климатическое исполнение
            list1['F33'].value = CableEntry  # Тип кабельного ввода
            # Взрывозащита
            exd_type('F35')  # Вид исполнения
            exd_entry('F37')  # Взрывозащищенность кабельного ввода

            list1['F36'].value = ExdEnv  # Вещество для определения степени взрывозащиты

            # Интерфейс
            list1['F39'].value = LED  # Наличие индикации
            list1['F39'].value = Screen  # Лок. интерфейс оператора для настройки датчика
            # Примечание
            list1['A42'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1

        if Designator == 'TG' or Designator == 'ТG':
            sh = openpyxl.load_workbook('TG.xlsx')
            list1 = sh.active
            # Общие данные
            list1['F16'].value = Designator + '-' + Position  # Позиция по ТХ
            list1['F17'].value = Type_device  # Измеряемый параметр
            list1['F18'].value = Environment  # Наименование измеряемой среды
            list1['F19'].value = Location  # Место установки
            # Технологические характеристики
            list1['F21'].value = Range  # Диапазон измерения
            list1['E21'].value = Unit #Единица измерения
            list1['F22'].value = Scale  # Шкала
            list1['E22'].value = Unit  # Единица измерения
            list1['F23'].value = Scale  # Абсолютная погрешность измерения
            list1['F24'].value = Pressure  # Давление среды
            list1['F25'].value = Temperature  # Температура в точке монтажа
            list1['F26'].value = ExplTemp  # Температура окружающего воздуха
            list1['F27'].value = Density  # Плотность
            list1['F28'].value = Viscosity  # Вязкость
            list1['F29'].value = Special_properties  # Особые свойства
            # Механические характеристики
            list1['F31'].value = Connection  # Соединение с процессом
            list1['F32'].value = Cartridge  # Гильза
            list1['F33'].value = ProcessMaterial  # Материальное исполнение в контакте с рабочей средой
            list1['F34'].value = IP  # Степень пылевлагозащиты IP
            list1['F35'].value = UHL  # Климатическое исполнение
            # Взрывозащита
            exd_type('F37')  # Вид исполнения

            list1['F38'].value = ExdEnv  # Вещество для определения степени взрывозащиты

            # Примечание
            list1['A40'].value = Note  # Примечание

            file_name = Designator + ' ' + Position + '.xlsx'
            saving(file_name)
            counter = counter + 1
print('Создание опросных листов завершено. Создано ', counter, ' опросных листов.')

input()
