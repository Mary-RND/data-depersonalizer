"""Примерочные тесты по критериям приёмки ТЗ (docs/technical-specification.md).

Каждая функция соответствует одному критерию AC-xx из раздела 12 ТЗ.
"""

import re

from depersonalizer.excel_core import depersonalize, restore
from depersonalizer.word_core import _replace_in_text


def _mapping_from_key(key_rows):
    mapping = {}
    for header, index, value in key_rows:
        mapping.setdefault(header, {})[index] = value
    return mapping


def _run(text, categories, custom_values=None):
    return _replace_in_text(text, categories, {}, {}, custom_values or [])


# AC-01: замена только выбранных колонок
def test_ac01_only_selected_columns_are_indexed():
    headers = ["ФИО", "Телефон", "Регион"]
    rows = [["Иванов Иван", "+7 900 111-22-33", "Москва"]] * 1000
    new_rows, key_rows, _ = depersonalize(headers, rows, {0, 1})
    assert re.fullmatch(r"ИД1_\d{4}", new_rows[0][0])
    assert re.fullmatch(r"ИД2_\d{4}", new_rows[0][1])
    assert new_rows[0][2] == "Москва"  # невыбранная колонка не меняется
    assert len(key_rows) == 2


# AC-02: одинаковые значения -> один индекс
def test_ac02_duplicates_map_to_one_index():
    headers = ["Телефон"]
    rows = [["+7 900 000-00-00"]] * 50
    _, key_rows, _ = depersonalize(headers, rows, {0})
    assert len(key_rows) == 1
    assert all(r[1] == "ИД1_0001" for r in key_rows)


# AC-03: обратная конвертация восстанавливает данные
def test_ac03_roundtrip_restores_all_values():
    headers = ["ФИО", "Телефон", "Регион"]
    rows = [
        ["Иванов Иван", "+7 900 111-22-33", "Москва"],
        ["Петров Пётр", "+7 900 222-33-44", "Тверь"],
    ]
    new_rows, key_rows, _ = depersonalize(headers, rows, {0, 1})
    mapping = _mapping_from_key(key_rows)
    assert restore(headers, new_rows, mapping) == rows


# AC-05: пустые ячейки не индексируются
def test_ac05_empty_cells_are_untouched():
    headers = ["ФИО"]
    rows = [["Иванов"], [""], [None], ["Петров"]]
    new_rows, key_rows, _ = depersonalize(headers, rows, {0})
    assert new_rows[1][0] == ""
    assert new_rows[2][0] is None
    assert len(key_rows) == 2


# AC-09: прямая конвертация Word заменяет данные и сохраняет токены
def test_ac09_word_masks_phone_email_fio():
    text = "Телефон +7 900 123-45-67, email ivanov@mail.ru, ФИО Иванов Иван Иванович"
    out = _run(text, {"Телефоны", "Email", "ФИО"})
    assert "ИДТЕЛ_0001" in out and "ИДЕМЕЙЛ_0001" in out and "ИДФИО_0001" in out
    for secret in ("+7 900 123-45-67", "ivanov@mail.ru", "Иванов Иван Иванович"):
        assert secret not in out


# AC-10: одно значение в тексте и таблице -> один индекс
def test_ac10_same_value_one_index():
    text = "В тексте +7 900 123-45-67, в таблице тоже +7 900 123-45-67"
    out = _run(text, {"Телефоны"})
    assert out.count("ИДТЕЛ_0001") == 2


# AC-13: адрес маскируется целиком
def test_ac13_address_masked_entirely():
    text = ("Юридический адрес: 428000, Чувашская Республика, г. Чебоксары, "
            "пр-т Мира, д. 19, офис 305.") + " Второй адрес: г. Москва, ул. Тверская, д. 1, стр. 2."
    out = _run(text, {"Адреса"})
    assert "ИДАДРЕС_0001" in out and "ИДАДРЕС_0002" in out
    for secret in ("428000", "Чувашская Республика", "Чебоксары", "Москва"):
        assert secret not in out


# AC-14: реквизиты и банковские реквизиты маскируются
def test_ac14_requisites_and_bank_masked():
    text = ("ИНН 7707083893, КПП 770701001, ОГРН 1027700132195, "
            "р/с 40702810900000001234, БИК 044525225, к/с 30101810400000000225")
    out = _run(text, {"Реквизиты организации", "Банковские реквизиты"})
    assert "ИДРЕКВ_" in out and "ИДБАНК_" in out
    for secret in ("7707083893", "770701001", "1027700132195",
                   "40702810900000001234", "044525225", "30101810400000000225"):
        assert secret not in out


# AC-15: телефон не «поедает» длинный счёт
def test_ac15_bank_account_not_eaten_by_phone():
    text = "р/с 40702810900000001234 и телефон 8 900 123 45 67"
    out = _run(text, {"Телефоны", "Банковские реквизиты"})
    assert "ИДБАНК_0001" in out and "ИДТЕЛ_0001" in out
    assert "40702810900000001234" not in out


# AC-16: паспортные данные маскируются, ИНН не затрагивается
def test_ac16_passport_masked_but_inn_untouched():
    text = "Паспорт: серия 45 12 № 345678. ИНН 7707083893."
    out = _run(text, {"Паспортные данные"})
    assert "ИДПАСП_0001" in out
    assert "45 12 № 345678" not in out
    assert "7707083893" in out