import pandas
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime


df_data = pd.read_excel('data.xlsx')
# удаление пустого столбца
df_data = df_data.drop('Unnamed: 5', axis=1)

date_row_index = []
# поиск индексов ячеек датафрейма, которые служат для разделения таблицы на месяца
for index, row in df_data[df_data['status'].str.contains(" 2021")].iterrows():
    date_row_index.append(index)
# разбивка изначального датафрейма, на датафрейми по месяцам
month_dataframe = {}
for list_index, df_index in enumerate(date_row_index):
    try:
        month_dataframe[df_data.loc[df_index]['status']] = df_data.iloc[date_row_index[list_index]:date_row_index[list_index+1]].drop(df_index, axis=0)
    except IndexError:
        month_dataframe[df_data.loc[df_index]['status']] = df_data.iloc[date_row_index[list_index]:].drop(df_index, axis=0)


# 1. Вычислите общую выручку за июль 2021 по тем сделкам, приход денежных
# средств которых не просрочен.

def count_paid_sum(df: pandas.DataFrame) -> float:
    paid_df = df[df.status != 'ПРОСРОЧЕНО']
    return round(paid_df['sum'].sum(), 2)


print('1.Выручка за июль 2021:', count_paid_sum(month_dataframe['Июль 2021']))


# 2. Как изменялась выручка компании за рассматриваемый период?
# Проиллюстрируйте графиком.

def profit_chart(month_dict: dict):
    profit = []
    months = []
    for key in month_dict:
        profit.append(round(count_paid_sum(month_dict[key])))
        months.append(key)

    plt.plot(months, profit, color='red', marker='o', linewidth=3)
    plt.ylabel('Прыбыль')
    plt.xlabel('Месяц')
    plt.title("График изменения прибыли")
    plt.show()
    return


profit_chart(month_dataframe)


# 3. Кто из менеджеров привлек для компании больше всего денежных средств в
# сентябре 2021?

def moust_profit_manager(df: pandas.DataFrame):
    manager_dict = {}
    for manager in df['sale'].unique():
        manager_dict[manager] = round(df[df.sale == manager]['sum'].sum())
    return max(manager_dict, key=manager_dict.get), manager_dict[max(manager_dict, key=manager_dict.get)]


sept_best_manager = moust_profit_manager(month_dataframe['Сентябрь 2021'])
print(f'3. Менеджер привлекший больше всего денежных средств в сентябре 2021 - {sept_best_manager[0]}, c суммой {sept_best_manager[1]} руб.')


# 4. Какой тип сделок (новая/текущая) был преобладающим в октябре 2021?
def moust_type(df: pandas.DataFrame):
    new_type = len(df[df.status != 'новая'].index)
    current_type = len(df[df.status != 'текущая'].index)
    if new_type > current_type:
        return 'Больше всего новых сделок'
    elif new_type < current_type:
        return 'Больше всего текущих сделок'

    return 'Сделок новая/текущая одинаковое колличество'


print('4. Наибольшее число сделок за октябрь 2021:', moust_type(month_dataframe['Октябрь 2021']))


# 5. Сколько оригиналов договора по майским сделкам было получено в июне 2021?
def count_original(df: pandas.DataFrame):
    originals = df[(df['receiving_date'] >= datetime.strptime('2021-05-01', '%Y-%m-%d')) &
                   (df['receiving_date'] < datetime.strptime('2021-06-01', '%Y-%m-%d')) &
                   (df['document'] == 'оригинал')]
    return len(originals)


print('5. Оригиналов договора по майским сделкам было получено в июне 2021:',count_original(month_dataframe['Июнь 2021']))


# За каждую заключенную сделку менеджер получает бонус, который рассчитывается
# следующим образом.
# 1) За новые сделки менеджер получает 7 % от суммы, при условии, что статус
# оплаты «ОПЛАЧЕНО», а также имеется оригинал подписанного договора с
# клиентом (в рассматриваемом месяце).
# 2) За текущие сделки менеджер получает 5 % от суммы, если она больше 10 тыс.,
# и 3 % от суммы, если меньше. При этом статус оплаты может быть любым,
# кроме «ПРОСРОЧЕНО», а также необходимо наличие оригинала подписанного
# договора с клиентом (в рассматриваемом месяце).
# Бонусы по сделкам, оригиналы для которых приходят позже рассматриваемого
# месяца, считаются остатком на следующий период, который выплачивается по мере
# прихода оригиналов. Вычислите остаток каждого из менеджеров на 01.07.2021.(Июнь)


def manager_salary_remainder(df: pandas.DataFrame, current_mounth=6):
    manager_salary_left = {}
    # убираем из датафрейма записи с типом ВНУТРЕННИЙ, так как они не входят в конечную выборку
    df = df[df['status'] != 'ВНУТРЕННИЙ']
    for manager in df['sale'].unique():
        # выделяем общие требования для расчета всех видов остатков, далее высчитываем отдельно пр
        base_manager_df = df[(df['receiving_date'] >= datetime.strptime(f'2021-0{current_mounth+1}-01', '%Y-%m-%d')) &
                             (df['receiving_date'] <= datetime.strptime(f'2021-0{current_mounth+1}-31', '%Y-%m-%d')) &
                             (df['document'] == 'оригинал') &
                             (df['sale'] == manager)]

        max_procent = base_manager_df[base_manager_df['new/current'] == 'новая']['sum'].sum() * 0.07
        mid_procent = base_manager_df[(base_manager_df['sum'] > 10000) & (base_manager_df['status'] != 'ПРОСРОЧЕНО')]['sum'].sum() * 0.05
        min_procent = base_manager_df[(base_manager_df['sum'] <= 10000) & (base_manager_df['status'] != 'ПРОСРОЧЕНО')]['sum'].sum() * 0.03

        manager_salary_left[manager] = round(max_procent + mid_procent + min_procent, 2)
    return manager_salary_left


print('6. Данные по остаткам каждого из менеджеров на 01.07.2021:\n', manager_salary_remainder(month_dataframe['Июнь 2021']))















