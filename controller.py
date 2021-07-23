import json


class Product():
    def __init__(self, name):
        self.f = File()
        self.name = name

    def checkProduct(self):
        if self.name in self.f.list['Продукты']:
            return True
        else:
            return False

    def addNewProduct(self):
        if self.checkProduct() == True:
            return True
        else:
            self.f.list['Продукты'].append(self.name)

    def delNewProduct(self):
        if self.checkProduct() == True:
            self.f.list['Продукты'].remove(self.name)
        else:
            print('Такого продукта нет!')


class Dish():
    def __init__(self, name, id):
        self.f = File()
        self.name = name
        self.id = str(id)

    def checkDish(self):
        if self.id in self.f.list['Блюда']:
            return True
        else:
            return False

    def addNewDish(self):
        if self.checkDish() == False:
            self.f.list['Блюда'][self.id] = {'Название': self.name, 'Продукты': {}}
            return False
        else:
            return True

    def addProductToDish(self, product, weight):
        try:
            self.f.list['Блюда'][self.id]['Продукты'][product] = weight
        except:
            print('something wrong')

    def delNewDish(self):
        if self.checkDish() == True:
            del self.f.list['Блюда'][self.id]
        else:
            print('Такого блюда нет!')


class Menu():
    def __init__(self, name, dateP):
        self.name = name
        self.date = dateP
        self.f = File()

    def addNewMenu(self):

        try:
            if self.name not in self.f.list['Меню'][self.date]:
                self.f.list['Меню'][self.date][self.name] = {}
        except:
            self.f.list['Меню'][self.date] = {self.name: {}}
            return True

    def copyMenu(self, newDate, newName):
        try:
            self.f.list['Меню'][newDate][newName] = self.f.list['Меню'][self.date][self.name]
        except:
            self.f.list['Меню'][newName] = {self.name: self.f.list['Меню'][self.date][self.name]}

    def delNewMenu(self):
        if self.checkNewMenu() == True:
            del self.f.list['Меню'][self.date][self.name]
            return True
        else:
            return False

    def checkNewMenu(self):
        if self.date in self.f.list['Меню']:
            if self.name in self.f.list['Меню'][self.date]:
                return True
            else:
                return False
        else:
            return False


class File():
    def __init__(self):
        try:
            with open('list.json', 'r') as file:
                self.list = json.load(file)
        except:
            with open('list.json', 'w') as file:
                self.list = {
                    'Продукты': [],
                    'Блюда': {},
                    'Меню': {}
                }
                json.dump(self.list, file)

    def saveFile(self):
        with open('list.json', 'w') as file:
            json.dump(self.list, file)
