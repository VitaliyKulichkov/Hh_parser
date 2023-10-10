from parser import Parser
from analyzer import Analyzer


def main():
    query = input('Введите поисковой запрос:')
    Parser().insert_to_excel(query)
    lst = Parser().parse_by_name_of_vacancy(query)
    Analyzer(lst=lst).draw_max_salary_histogram(query)
    Analyzer(lst=lst).draw_min_salary_histogram(query)
    Analyzer(lst=lst).draw_average_salary_histogram(query)


if __name__ == '__main__':
    main()
