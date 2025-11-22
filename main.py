from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas
from collections import defaultdict


def get_year_word(n: int) -> str:
    if 11 <= n % 100 <= 14:
        return "лет"
    last_digit = n % 10
    if last_digit == 1:
        return "год"
    elif 2 <= last_digit <= 4:
        return "года"
    else:
        return "лет"


def parse_excel(filename: str) -> list:
    excel_data_df = pandas.read_excel(filename, sheet_name='Лист1', na_values=" ", keep_default_na=False)
    records = excel_data_df.to_dict(orient='records')
    wines_by_category = defaultdict(list)

    for record in records:
        category = record.get("Категория", "").strip()
        wines_by_category[category].append(record)
    return wines_by_category


def main():
    start_year = 1920
    current_year = datetime.now().year
    years = current_year - start_year

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        year_text=f"Уже {years} {get_year_word(years)} с вами",
        wines=parse_excel("wine3.xlsx"),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
