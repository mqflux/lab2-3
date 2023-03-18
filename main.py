import pandas as pd
from random import choice, randint, uniform
from openpyxl import load_workbook


class Item:
    def __init__(self, name, group, n_sold, sell_price, buy_price, discount):
        self.name = name
        self.group = group
        self.n_sold = n_sold
        self.sell_price = sell_price
        self.buy_price = buy_price
        self.discount = discount
        self.profit = sell_price * n_sold - buy_price * n_sold


def randomize_item():
    group_dict = {
        "food": ["Мясо", "Мёд", "Мука", "Сыр", "Молоко", "Масло", "Рыба"],
        "furniture": ["Стул", "Кресло", "Диван", "Стол", "Полка", "Кровать", "Телевизор", "Шкаф"],
        "entertainment": ["Билет в кино", "Билет на концерт", "Книга", "Билет в парк развлечений"],
        "materials": ["Бетон", "Плитка", "Доска", "Гвозди", "Гайки"],
    }
    group = choice(list(group_dict.keys()))
    name = choice(group_dict[group])

    n_sold = randint(1, 100)
    sell_price = randint(250, 10000)
    buy_price = sell_price * uniform(0.33, 0.93)

    if randint(0, 100) > 80:
        discount = randint(10, 25)
    else:
        discount = 0

    return Item(name, group, n_sold, sell_price, buy_price, discount)


def generate_data(n):
    name_list, group_list, n_sold_list, sell_price_list, \
        buy_price_list, discount_list, profit_list = [], [], [], [], [], [], []

    items = [randomize_item() for _ in range(n)]
    items.sort(key=lambda x: x.group)

    for item in items:

        name_list.append(item.name)
        group_list.append(item.group)
        n_sold_list.append(item.n_sold)
        sell_price_list.append(item.sell_price)
        buy_price_list.append(item.buy_price)
        discount_list.append(item.discount)
        profit_list.append(item.profit)

    df = pd.DataFrame({"Group": group_list,
                       "Name": name_list,
                       "Amount": n_sold_list,
                       "Sell Price": sell_price_list,
                       "Buy Price": buy_price_list,
                       "Discount": discount_list,
                       "Profit": profit_list})

    df.to_json("newExcelData.json")


def update_data(source_file, out_file):
    df_new = pd.read_json(source_file)
    try:
        df_old = pd.read_excel(out_file, sheet_name="Raw")
    except Exception:
        df_old = pd.DataFrame()
    writer = pd.ExcelWriter(out_file, engine='openpyxl')

    try:
        df_old = df_old.append(df_new)
        df_old.to_excel(writer, sheet_name="Raw", index=False)
    except Exception:
        df_new.to_excel(writer, sheet_name="Raw", index=False)

    if df_old.empty:
        update_sum_sheet(df_old, writer)
    else:
        update_sum_sheet(df_new, writer)

    writer.close()


def update_sum_sheet(data_frame, writer):
    group_name, profit_sum = [], []

    for i in range(len(data_frame["Group"])):
        if data_frame["Group"][i] in group_name:
            index = group_name.index(data_frame["Group"][i])
            profit_sum[index] += data_frame["Profit"][i]
        else:
            group_name.append(data_frame["Group"][i])
            profit_sum.append(data_frame["Profit"][i])

    out_frame = pd.DataFrame({
        "Group": group_name,
        "Profit": profit_sum
    })

    out_frame.to_excel(writer, sheet_name="Sum", header=False, index=False)


if __name__ == "__main__":
    generate_data(444)
    update_data("NewExcelData.json", "newExcelData.xlsx")
