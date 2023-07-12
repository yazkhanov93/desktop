import json
from tkinter import *
from tkinter import messagebox as ms, filedialog
import sqlite3
from tkinter import ttk
import pandas as pd
from pandas import *
import openpyxl

import requests

with sqlite3.connect('quit.db') as db:
    c = db.cursor()

c.execute('CREATE TABLE IF NOT EXISTS user (username TEXT NOT NULL PRIMARY KEY,password TEX NOT NULL);')
db.commit()
db.close()

# urLink = 'http://127.0.0.1:8000/' 
urLink = "https://km-aparts.com/"
# main Class
class main:
    def __init__(self, master):
        # Window
        self.master = master
        # Some Usefull variables
        self.username = StringVar()
        self.password = StringVar()
        # Create Widgets
        self.widgets()

    def openfile1(self):
        self.filename1 = filedialog.askopenfilename(initialdir='', title='Выберите Файл')
        self.label = ttk.Label()
        self.label.grid(column=1, row=6)
        self.label.configure(text=self.filename1.split("/")[-1])

    def uploadProduct(self):
        try:
            df = pd.read_excel(self.filename1)
            df = df[["Код",'Наименование', 'Дополнительное Наименование', 'Модель',
            'Год', 'Мотор', 'Размер', 'Компания', 'Made_IN',
            'Made_IN на экран', 'Цена _Dollar', 'vip регистрац', 'регистрац',
            'Остаток_кол-во', 'Company Part Number',
            'Original Part Number', 'Вес_кг', 'Кузов', 'ЕД', 'ОПТ_Цена _Dollar','main','Серийный_№']]
            df = df.dropna(how="all")
            df = df.fillna("")
            df = df.rename(columns={"Код":'code','Наименование':'title', 'Дополнительное Наименование':'description', 'Модель':'model',
            'Год':'years', 'Мотор':'motor', 'Размер':'size', 'Компания':'company_name', 'Made_IN':'made_in',
            'Made_IN на экран':'made_in_visible', 'Цена _Dollar':'price', 'vip регистрац':'is_visible_vip', 'регистрац':'is_visible_all',
            'Остаток_кол-во':'quantity', 'Company Part Number':'company_part_number',
            'Original Part Number':'original_part_number', 'Вес_кг':'weight', 'Кузов':'cascade', 'ЕД':'unit', 'ОПТ_Цена _Dollar':'wholeSalePrice','Серийный_№':"serial"})
            df["quantity"] = df["quantity"].replace("", 0)
            df["main"] = df["main"].replace("", 0)
        
            df = df.astype({"code":"str","title":"str","description":"str","model":"str","years":"str","motor":"str","size":"str","company_name":"str", "company_part_number":"str", "original_part_number":"str","weight":"str","cascade":"str" , "quantity":"int","main":"int"})
            
            df["serial"] = df["serial"].str.replace("-","")
            df["serial"] = df["serial"].fillna("")
            df["serial"] = df["serial"].str.upper()

            # df["years"] = df["years"].str.replace("",0)
            

            df["company_part_number"] = df["company_part_number"].str.replace("-| |:|#|;|$|_","")
            df["company_part_number"] = df["company_part_number"].str.upper()
            
            df["original_part_number"] = df["original_part_number"].str.replace("-| |:|#|;|$|_","")
            df["original_part_number"] = df["original_part_number"].str.upper()

            df["size"] = df["size"].str.replace("-|/|:|#|_","*")

            df["made_in_visible"] = df["made_in_visible"].replace(1, True)
            df["made_in_visible"] = df["made_in_visible"].replace(0, False)
            df["made_in_visible"] = df["made_in_visible"].replace("", False)

            df["is_visible_all"] = df["is_visible_all"].replace(1, True)
            df["is_visible_all"] = df["is_visible_all"].replace(0, False)
            df["is_visible_all"] = df["is_visible_all"].replace("", False)

            df["is_visible_vip"] = df["is_visible_vip"].replace(1, True)
            df["is_visible_vip"] = df["is_visible_vip"].replace(0, False)
            df["is_visible_vip"] = df["is_visible_vip"].replace("", False)

            df["motor"] = df["motor"].str.replace("  ", " ")
            # print(df)
            df = df.to_dict(orient="records")
            for i in df:
                if i["years"].isdigit() and i["years"].__contains__("-"):
                    year_start =int(i["years"].split("-")[0])
                    year_end = int(i["years"].split("-")[-1])
                elif i["years"].isdigit():
                    year_start = int(i["years"])
                    year_end = int(i["years"])
                else:
                    year_start = 0
                    year_end = 0
                i.update({"year_start":year_start, "year_end":year_end})
                del i["years"]
            
            i = 0
            dt = [] 
            while i < len(df):
                dt.append(df[i:i+100])
                i+=100
                for j in dt:
                    json_data = json.dumps(j, ensure_ascii=True)
                
                upload = requests.post(url=url + "/upload-product/", data=json_data, headers={"Content-Type":"application/json; charset=utf-8"})
                # update = requests.put(url=url + "/upload-product/", data=json_data, headers={"Content-Type":"application/json; charset=utf-8"})
            # print(len(df))
        except Exception as e:
            ms.showerror(title="Ошибка", message=e)
    
    def uploadComp(self):
        try:
            df = pd.read_excel(self.filename2)
            df = df[["наименование","партномер_а", "модель_а","год_а", 'мотор_а','кузов_а', 'код','дополнительно']]
            df = df.dropna(how="all")
            df = df.fillna("")
            df = df.rename(columns={"наименование":"title","партномер_а":"original_part_number", "модель_а":"model","год_а":"years","мотор_а":"motor","кузов_а":"cascade","код":"product","дополнительно":"description"})
            df = df.astype({"title":"str","original_part_number":"str","model":"str","motor":"str","years":"str","cascade":"str","product":"str"})
            df = df.to_dict(orient="records")
            for i in df:
                if i["years"].isdigit() or i["years"].__contains__("-"):
                    year_start = i["years"].split("-")[0]
                    year_end = i["years"].split("-")[-1]
                year_start = int(year_start)
                year_end = int(year_end)
                i.update({"year_start":year_start, "year_end":year_end})
                del i["years"]
            comp = requests.get(url=url + "/upload-comp/")
            comp = comp.json()
            for i in comp:
                del i["id"]
            
            dt = []
            for i in df:
                if i in comp:
                    pass
                else:
                    dt.append(i)
            i = 0
            c=[]
            while i < len(dt):
                c.append(dt[i:i+100])
                i += 100
                for j in c:
                    json_data = json.dumps(j, ensure_ascii=True)
                # print("Go...")
                # print(json_data)
                upload = requests.post(url=url + "/upload-comp/", data=json_data, headers={"Content-Type":"application/json; charset=utf-8"})
        except Exception as e:
            ms.showerror(title="Ошибка", message=e)

    
    def openfile2(self):
        self.filename2 = filedialog.askopenfilename(initialdir='', title='Выберите Файл')
        self.label = ttk.Label()
        self.label.grid(column=2, row=6)
        self.label.configure(text=self.filename2.split("/")[-1])

    def uploaddubai(self):
        try:
            df = pd.read_excel(self.filename3)
            df = df[["Код",'Наименование', 'Дополнительное Наименование', 'Модель',
            'Год', 'Мотор', 'Размер', 'Компания', 'Made_IN',
            'Made_IN на экран', 'Цена _Dollar', 'vip регистрац', 'регистрац',
            'Остаток_кол-во', 'Company Part Number',
            'Original Part Number', 'Вес_кг', 'Кузов', 'ЕД', 'ОПТ_Цена _Dollar','main','Серийный_№']]
            df = df.dropna(how="all")
            df = df.fillna("")
            df = df.rename(columns={"Код":'code','Наименование':'title', 'Дополнительное Наименование':'description', 'Модель':'model',
            'Год':'years', 'Мотор':'motor', 'Размер':'size', 'Компания':'company_name', 'Made_IN':'made_in',
            'Made_IN на экран':'made_in_visible', 'Цена _Dollar':'price', 'vip регистрац':'is_visible_vip', 'регистрац':'is_visible_all',
            'Остаток_кол-во':'quantity', 'Company Part Number':'company_part_number',
            'Original Part Number':'original_part_number', 'Вес_кг':'weight', 'Кузов':'cascade', 'ЕД':'unit', 'ОПТ_Цена _Dollar':'wholeSalePrice','Серийный_№':"serial"})
            df["quantity"] = df["quantity"].replace("", 0)
            df["main"] = df["main"].replace("", 0)
        
            df = df.astype({"code":"str","title":"str","description":"str","model":"str","years":"str","motor":"str","size":"str","company_name":"str", "company_part_number":"str", "original_part_number":"str","weight":"str","cascade":"str" , "quantity":"int","main":"int"})
            
            df["serial"] = df["serial"].str.replace("-","")
            df["serial"] = df["serial"].fillna("")
            df["serial"] = df["serial"].str.upper()

            # df["years"] = df["years"].str.replace("",0)
            

            df["company_part_number"] = df["company_part_number"].str.replace("-| |:|#|;|$|_","")
            df["company_part_number"] = df["company_part_number"].str.upper()
            
            df["original_part_number"] = df["original_part_number"].str.replace("-| |:|#|;|$|_","")
            df["original_part_number"] = df["original_part_number"].str.upper()

            df["size"] = df["size"].str.replace("-|/|:|#|_","*")

            df["made_in_visible"] = df["made_in_visible"].replace(1, True)
            df["made_in_visible"] = df["made_in_visible"].replace(0, False)
            df["made_in_visible"] = df["made_in_visible"].replace("", False)

            df["is_visible_all"] = df["is_visible_all"].replace(1, True)
            df["is_visible_all"] = df["is_visible_all"].replace(0, False)
            df["is_visible_all"] = df["is_visible_all"].replace("", False)

            df["is_visible_vip"] = df["is_visible_vip"].replace(1, True)
            df["is_visible_vip"] = df["is_visible_vip"].replace(0, False)
            df["is_visible_vip"] = df["is_visible_vip"].replace("", False)

            df["motor"] = df["motor"].str.replace("  ", " ")
            # print(df)
            df = df.to_dict(orient="records")
            for i in df:
                if i["years"].isdigit() and i["years"].__contains__("-"):
                    year_start =int(i["years"].split("-")[0])
                    year_end = int(i["years"].split("-")[-1])
                elif i["years"].isdigit():
                    year_start = int(i["years"])
                    year_end = int(i["years"])
                else:
                    year_start = 0
                    year_end = 0
                i.update({"year_start":year_start, "year_end":year_end})
                del i["years"]
            
            i = 0
            dt = [] 
            while i < len(df):
                dt.append(df[i:i+100])
                i+=100
                for j in dt:
                    json_data = json.dumps(j, ensure_ascii=True)
                
                # upload = requests.post(url=url + "/upload-dubai/", data=json_data, headers={"Content-Type":"application/json; charset=utf-8"})
                # update = requests.put(url=url + "/upload-product/", data=json_data, headers={"Content-Type":"application/json; charset=utf-8"})
            print(len(df))
        except Exception as e:
            ms.showerror(title="Ошибка", message=e)

    def openfile3(self):
        self.filename3 = filedialog.askopenfilename(initialdir='', title='Выберите Файл')
        self.label = Label()
        self.label.grid(column=3, row=6)
        self.label.configure(text=self.filename3.split("/")[-1])


    # Login Function
    def login(self):
        # Establish Connection
        with sqlite3.connect('quit.db') as db:
            c = db.cursor()

        # Find user If there is any take proper action
        find_user = ('SELECT * FROM user WHERE username = ? and password = ?')
        c.execute(find_user, [(self.username.get()), (self.password.get())])
        result = c.fetchall()
        month = [
            "Январь",
            "Февраль",
            "Март",
            "Апрель",
            "Май",
            "Июнь",
            "Июль",
            "Август",
            "Сентябр",
            "Октябр",
            "Ноябр",
            "Декабр"
        ]
        self.month_numbers = {
            "Январь": "01",
            "Февраль": "02",
            "Март": "03",
            "Апрель": "04",
            "Май": "05",
            "Июнь": "06",
            "Июль": "07",
            "Август": "08",
            "Сентябр": "09",
            "Октябр": "10",
            "Ноябр": "11",
            "Декабр": "12"
        }
        order_status = ["В_ПРОЦЕССЕ", "ПРИНЯТ", "ОТПРАВЛЕН", "ДОСТАВЛЕН", "ОТМЕНЕН"]
        if result:
            self.logf.grid_forget()
            self.label = Label(text="Загрузить Товар", font=("Arial, 16")).grid(column=1, row=3, padx=50, pady=10)
            self.button1 = Button(text="Выберите Файл", command=self.openfile1, font=("Arial, 11"), width=12).grid(
                column=1, row=4, padx=10, pady=30)
            self.button1_1 = Button(text="Загрузить", font=("Arial, 11"), width=12, command=self.uploadProduct).grid(column=1,
                                                                                                               row=5,
                                                                                                               padx=10,
                                                                                                               pady=10)

            self.label = Label(text="Загрузить Совместимых", font=("Arial, 16")).grid(column=2, row=3, padx=30, pady=30)
            self.button2 = Button(text="Выберите Файл", command=self.openfile2, font=("Arial, 11"), width=12).grid(
                column=2, row=4, padx=10, pady=30)
            self.button2_2 = Button(text="Загрузить", font=("Arial, 11"), width=12, command=self.uploadComp).grid(column=2,
                                                                                                               row=5,
                                                                                                               padx=10,
                                                                                                               pady=10)
            self.label = Label(text="Загрузить Дубай", font=("Arial, 16")).grid(column=3, row=3, padx=30, pady=30)
            self.button3 = Button(text="Выберите Файл", command=self.openfile3, font=("Arial, 11"), width=12).grid(
                column=3, row=4, padx=10, pady=30)
            self.button3_3 = Button(text="Загрузить", font=("Arial, 11"), width=12, command=self.uploaddubai).grid(column=3,
                                                                                                               row=5,
                                                                                                               padx=10,
                                                                                                               pady=10)
            self.label = Label(text="Скачать Список Заказов", font="Arial, 16").grid(column=1, row=7, pady=50)
            self.label = Label(text="Год", font=("Arial, 12")).grid(column=1, row=8)
            self.year = Entry(width=6, font=("", 16))
            self.year.insert(0,"2022")
            self.year.grid(column=1, row=9, pady=10)

            self.label = Label(text="Месяц",font=("Arial, 12")).grid(column=2, row=8)
            self.month = ttk.Combobox(values=month, state="readonly", font=("Arial, 16"))
            self.month.set("Январь")
            self.month.grid(column=2, row=9, pady=10)

            self.label = Label(text="Статус заказа",font=("Arial, 12")).grid(column=3, row=8)
            self.status = ttk.Combobox(values=order_status, state="readonly",font=("Arial, 14"), width=15)
            self.status.set("ДОСТАВЛЕН")
            self.status.grid(column=3, row=9)

            self.button = Button(text="Скачать", command=self.downloadOrders, font=("Arial, 14"))
            self.button.grid(column=4, row=9, padx=5)

            # self.head['text'] =self.openfile1.filename1
        else:
            ms.showerror('Oops!', 'Имя или пароль не совпадают.')

    def downloadOrders(self):
        try:
            self.year_from = self.year.get()
            self.month_from = self.month.get()
            self.status2 = self.status.get()
            date = str(self.year_from)+"-"+str(self.month_numbers[self.month_from])
            xlfilename = ("Список_Заказов_" + self.status2 + "_" + date + ".xlsx")
            orders_list = requests.get(url=urLink + "api/order/order-list/", timeout=60)
            orders_list = orders_list.json()
            product_list = requests.get(url=urLink + "api/products/product-post/", timeout=60)
            product_list = product_list.json()
            part_model = requests.get(url=urLink + "api/products/part-model-list/", timeout=60)
            part_model = part_model.json()
            customers = requests.get(url=urLink + "api/user/users-list/", timeout=60)
            customers = customers.json()
            order_list = []
            for i in orders_list:
                if i["order_status"]:
                    if i["order_status"] == self.status2:
                        if i["created_at"][0:7] == date:
                            order_list.append(i)
            for i in order_list:
                for j in customers:
                    if i["customer_id"] == j["id"]:
                        i["customer_id"] = j["username"]
            for i in order_list:
                i.update({"Order Number": i["id"], "Customer": i["customer_id"]})
                del i["customer_id"], i["id"]
            order = []
            for i in order_list:
                order.append(i)
                for j in i["products"]:
                    order.append(j)
                del i["products"]
            for i in order:
                if "id" in i:
                    del i["id"]
                if "order_id" in i:
                    del i["order_id"]
                if "created_at" in i:
                    i["created_at"] = i["created_at"][:10]
                if "updated_at" in i:
                    i["updated_at"] = i["updated_at"][:10]

            order_list = pd.DataFrame(order)
            order = order_list.rename(
                columns={"created_at": "Дата заказа", "updated_at": "Обработка заказа", "order_status":"Статус Заказа","Order Number": "Номер заказа",
                         "Customer": "Номер клиента", "quantity": "Количество", "price": "Цена",
                         "product_id": "Код Продукта"})
            order = order.fillna("0")
            order = order.to_dict(orient="records")
            for i in order:
                for j in product_list:
                    if i["Код Продукта"] == j["code"]:
                        i.update({"Названия": j["title"], "Модель": j["part_model"], "Company Part Number": j["company_pn"],
                                  "Original Part Number": j["original_pn"]})
            for i in order:
                if i["Дата заказа"] == "0":
                    i["Дата заказа"] = ""
                if i["Обработка заказа"] == "0":
                    i["Обработка заказа"] = ""
                if i["Статус Заказа"] == "0":
                    i["Статус Заказа"] = ""
                if i["Номер заказа"] == "0":
                    i["Номер заказа"] = ""
                else:
                    i["Номер заказа"] = int(i["Номер заказа"])
                if i["Номер клиента"] == "0":
                    i["Номер клиента"] = ""
                if i["Количество"] == "0":
                    i["Количество"] = ""
                else:
                    i["Количество"] = int(i["Количество"])
                if i["Цена"] == "0":
                    i["Цена"] = ""
                if i["Код Продукта"] == "0":
                    i["Код Продукта"] = ""
            for i in order:
                for j in part_model:
                    if "Модель" in i:
                        if i["Модель"] == j["id"]:
                            i["Модель"] = j["title"]
                            i.update({"year_start": j["year_start"], "year_end": j["year_end"], "Мотор": j["motor"]})
            for i in order:
                if "year_start" in i:
                    if "year_end" in i:
                        years = str(i["year_start"]) + "-" + str(i["year_end"])
                        i.update({"Год": years})
                        del i["year_start"], i["year_end"]
            if len(order) > 0:
                order = pd.DataFrame(order)
                order = order.fillna("")
                # print(xlfilename)
                order = order.to_excel(xlfilename, sheet_name="orders", index=False)
                ms.showinfo(title="Cохранено как ", message=xlfilename)
            else:
                ms.showinfo(message="Нет заказов")
        except:
            ms.showwarning(message="проверьте подключение к интернету")


    def log(self):
        self.username.set('')
        self.password.set('')
        self.logf.grid()

    # Draw Widgets
    def widgets(self):
        self.logf = Frame(self.master, padx=80, pady=10)
        # self.logf.config(bg="#008080")
        Label(self.logf, text='Username: ', font=('', 20), pady=50, padx=50).grid(column=2, row=3)
        Entry(self.logf, textvariable=self.username, bd=5, font=('', 15)).grid(row=3, column=3,pady=50, padx=50)
        Label(self.logf, text='Password: ', font=('', 20), pady=5, padx=5).grid(column=2, row=4)
        Entry(self.logf, textvariable=self.password, bd=5, font=('', 15), show='*').grid(row=4, column=3)
        Button(self.logf, text=' Login ', bd=3, font=('', 12), padx=5, pady=5, command=self.login).grid(row=5, column=3)

        self.logf.grid()


if __name__ == '__main__':
    # Create Object
    # and setup window
    root = Tk()
    root.title('KM-Parts')
    # root.config(bg="#008080")
    root.geometry('1000x500')
    main(root)
    root.mainloop()