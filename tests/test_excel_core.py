from depersonalizer.excel_core import depersonalize, restore


HEADERS = ["ФИО", "Телефон", "Регион"]
ROWS = [
    ["Иванов Иван", "+7 900 111-22-33", "Москва"],
    ["Петров Пётр", "+7 900 111-22-33", "Москва"],
    ["Сидорова Анна", "+7 900 444-55-66", ""],
    ["", "", "Тверь"],
]


def _mapping_from_key(key_rows):
    mapping = {}
    for header, index, value in key_rows:
        mapping.setdefault(header, {})[index] = value
    return mapping


def test_index_format():
    _, key_rows, _ = depersonalize(HEADERS, ROWS, {0, 1})
    indexes = {row[1] for row in key_rows}
    assert "ИД1_0001" in indexes
    assert "ИД2_0001" in indexes


def test_duplicates_get_one_index():
    new_rows, key_rows, _ = depersonalize(HEADERS, ROWS, {1})
    assert new_rows[0][1] == new_rows[1][1]
    assert len(key_rows) == 2  # два уникальных телефона


def test_empty_cells_not_replaced_and_not_in_key():
    new_rows, key_rows, _ = depersonalize(HEADERS, ROWS, {0})
    assert new_rows[3][0] == ""
    assert all(row[2] != "" for row in key_rows)


def test_roundtrip_restore():
    new_rows, key_rows, _ = depersonalize(HEADERS, ROWS, {0, 1})
    mapping = _mapping_from_key(key_rows)
    restored = restore(HEADERS, new_rows, mapping)
    assert restored == ROWS


def test_restore_ignores_unselected_column():
    new_rows, key_rows, _ = depersonalize(HEADERS, ROWS, {0})
    mapping = _mapping_from_key(key_rows)
    restored = restore(HEADERS, new_rows, mapping)
    assert restored == ROWS