from depersonalizer.word_core import _replace_in_text


def _run(text, categories, custom_values=None):
    return _replace_in_text(text, categories, {}, {}, custom_values or [])


def test_phone_replaced():
    out = _run("Звоните +7 900 123-45-67 сегодня", {"Телефоны"})
    assert "ИДТЕЛ_0001" in out
    assert "+7 900 123-45-67" not in out


def test_email_replaced():
    out = _run("Пишите ivanov@mail.ru, пожалуйста", {"Email"})
    assert "ИДЕМЕЙЛ_0001" in out
    assert "ivanov@mail.ru" not in out


def test_fio_replaced():
    out = _run("Встретились с Иванов Иван Иванович в офисе", {"ФИО"})
    assert "ИДФИО_0001" in out
    assert "Иванов Иван Иванович" not in out


def test_custom_values_replaced():
    out = _run("Компания «ООО Ромашка» выиграла тендер", {"Телефоны"}, ["ООО Ромашка"])
    assert "ИДДОП_0001" in out
    assert "ООО Ромашка" not in out


def test_req_masked():
    out = _run("ИНН 7707083893, КПП 770701001", {"Реквизиты организации"})
    assert "ИДРЕКВ_0001" in out
    assert "7707083893" not in out
    assert "770701001" not in out


def test_bank_masked():
    out = _run("р/с 40702810900000001234, БИК 044525225", {"Банковские реквизиты"})
    assert "ИДБАНК_0001" in out
    assert "40702810900000001234" not in out
    assert "ИДБАНК_0002" in out


def test_bank_account_not_partially_masked_as_phone():
    text = "р/с 40702810900000001234 и телефон 8 900 123 45 67"
    out = _run(text, {"Телефоны", "Банковские реквизиты"})
    assert "40702810900000001234" not in out
    assert "8 900 123 45 67" not in out
    assert "ИДБАНК_0001" in out
    assert "ИДТЕЛ_0001" in out


def test_address_masked_entirely():
    text = "428000, Чувашская Республика, г. Чебоксары, пр-т Мира, д. 19, офис 305."
    out = _run(text, {"Адреса"})
    assert "ИДАДРЕС_0001" in out
    assert "428000" not in out
    assert "Чебоксары" not in out
    assert "пр-т Мира" not in out


def test_passport_masked():
    out = _run("Паспорт: серия 45 12 № 345678.", {"Паспортные данные"})
    assert "ИДПАСП_0001" in out
    assert "45 12 № 345678" not in out


def test_passport_does_not_touch_inn():
    out = _run("Паспорт: серия 45 12 № 345678. ИНН 7707083893.", {"Паспортные данные"})
    assert "ИДПАСП_0001" in out
    assert "7707083893" in out