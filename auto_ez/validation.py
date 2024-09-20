import re


def valid(data):
    add_alert = []
    if len(data["bs"]) > 40:
        add_alert.append('Очень длинное название БС!')
    if len(data["bs2"]) > 40:
        add_alert.append('Очень длинное название БС!')
    if isinstance(data["num_protocol"], int):
        add_alert.append('Номер протокола не число!')
    if not data["date_protocol"]:
        add_alert.append('Нет даты протокола!')
    if not data["date_protocol"]:
        add_alert.append('Нет даты измерений!')
    if not data["source"]:
        add_alert.append('Нет источника излучения!')
    if not data["result1"] or not data["result0"]:
        add_alert.append('Нет таблицы с результатами!')
    for dvs in data["devices"]:
        if not dvs[0]:
            add_alert.append('Нет названия прибора!')
        elif not dvs[1]:
            add_alert.append('Нет заводского номера прибора!')
        elif not dvs[2]:
            add_alert.append('Нет номера сертификата прибора!')
        elif not dvs[3]:
            add_alert.append('Нет даты сертификата прибора!')
    if not re.search(r'\+', data["temp"]) and not re.search(r'-', data["temp"]):
        add_alert.append('Показатель температуры некорректный!')
    if not re.search(r'\d\d', data["humidity"]):
        add_alert.append('Показатель влажности некорректный!')
    if not re.search(r'\d\d\d', data["pressure"]):
        add_alert.append('Показатель атм. давления некорректный!')
    if not re.search(r'\D{5,25}', data["repres"]):
        add_alert.append('ФИО представителя некорректное!')
    if not re.search(r'\D{5,25}', data["specialist"]):
        add_alert.append('ФИО специалиста некорректное!')
    return add_alert
