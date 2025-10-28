TABLE_CONFIG = {
    "parts": {
        "title": "Спецификация",
        "headers": ["Номенклатура", "Код", "Кол-во"],
        "fields": ["name", "code", "quantity"],
        "col_widths": [100, 35, 20],
        "row_height": 12,
        "color": "#4CAF50",
        "wrap": [0]
    },
    "operations": {
        "title": "Операции",
        "headers": ["№", "Код", "Наименование", "Оборудование", "Tподг", "Tшт"],
        "fields": ["number", "code", "name", "equipment", "prep_time", "unit_time"],
        "col_widths": [15, 25, 60, 50, 20, 20],
        "row_height": 14,
        "color": "#2196F3",
        "wrap": [2, 3]
    },
    "workshops": {
        "title": "Расцеховка",
        "headers": ["Цех", "Участок", "РМ"],
        "fields": ["workshop", "section", "workplace"],
        "col_widths": [60, 60, 60],
        "row_height": 10,
        "color": "#FF9800",
        "wrap": []
    },
    "equipment": {
        "title": "Оборудование",
        "headers": ["Наименование", "Артикул", "Примечание"],
        "fields": ["name", "article", "note"],
        "col_widths": [70, 50, 60],
        "row_height": 12,
        "color": "#9C27B0",
        "wrap": [0, 2]
    }
}

DETAIL_FIELDS = [
    ("Организация", "organization"),
    ("Обозначение изделия", "product_code"),
    ("Обозначение документа", "document_code"),
    ("Разработал", "developer"),
    ("Проверил", "checker")
]
