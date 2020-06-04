import PySimpleGUI as sg
import xlrd
import math
import datetime
import psycopg2 as pg 

from psycopg2.extras import DictCursor


##########################################	FUNCTIONS TO GET DATA ########################################
## connection to DB and queries (check_* get_* add_* update_* functions)

def dict_to_list(d):
    l = []
    for val in d.values():
        l.append(val)
    return l

def get_last_id():
	print('get_last_id')
	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select get_last_id()")
	row = cursor.fetchone()
	cursor.close()
	return row[0]

def check_id(id):
	print('check_id')
	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select check_id(%s)", (id, ))
	row = cursor.fetchone()
	cursor.close()
	return row[0]	

def get_status_with_date(id):
	print('get_status_with_date')
	if check_id(id) == 'no':
		return 'no', 'no'

	
	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select get_status(%s)", (id, ))
	row = cursor.fetchone()
	if row[0] == 'no':
		cursor.close()
		return 'no', 'no'
	status = row[0]
	
	cursor.execute("select get_status_date(%s, %s)", (id, status))
	row = cursor.fetchone()
	if row[0] is None:
		cursor.close()
		return 'no', 'no'
	date = row[0]

	cursor.close()
	return status, str(date)


# def add_pre_order(order_data, pay_data, label):
# 	pass
# 	print('add_pre_order')
# 	print(order_data, pay_data, label)
# 	return 'ok'
# 	#res = db(select add_pre_order)

def add_order(order_data, pay_data):
	print('add_order')
	print(order_data, pay_data)

	delivery = order_data['msq'] * 180
	debt = float(order_data['price']) - float(pay_data['prepay'])
	date_date = datetime.datetime.strptime(order_data['date'], '%Y-%m-%d').date()

	cursor = conn.cursor(cursor_factory=DictCursor)

	cursor.execute("select add_ord_main_info(%s, %s, %s, %s)", (int(order_data['id']), order_data['client'], order_data['person'], date_date))
	row = cursor.fetchone()
	if row[0] == 'no':
		print('ord_main_info failed')
		cursor.close()
		return 'no'

	cursor.execute("select add_ord_content(%s, %s, %s, %s, %s, %s, %s)", (int(order_data['id']), order_data['color'], order_data['type'], order_data['shape'], order_data['edge'], order_data['msq'], int(order_data['items'])))
	row = cursor.fetchone()
	if row[0] == 'no':
		print('ord_content failed')
		cursor.close()
		return 'no'

	cursor.execute("select add_ord_money(%s, %s, %s, %s, %s, %s, %s, %s)", (int(order_data['id']), float(order_data['price']), float(pay_data['prepay']), debt, pay_data['type'], pay_data['doc'], pay_data['date'], delivery))
	row = cursor.fetchone()
	if row[0] == 'no':
		print('ord_money failed')
		cursor.close()
		return 'no'

	conn.commit()
	cursor.close()

	return 'ok'


def update_status(new_status, new_date, id):
	print('update_status')
	print(new_status, new_date, id)

	if check_id(id) == 'no':
		return 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select update_status(%s, %s, %s)", (id, new_status, new_date))
	row = cursor.fetchone()
	if row[0] == 'ok':
		conn.commit()
	cursor.close()
	return row[0]	
	

def add_prepay(id, prepay):

	print('add_prepay')
	print(id, prepay)

	if check_id(id) == 'no':
		return 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select add_prepay(%s, %s)", (int(id), float(prepay)))
	row = cursor.fetchone()
	if row[0] == 'ok':
		conn.commit()
	cursor.close()
	return row[0]	
	

def set_bill(id, bill):

	print('set_bill')
	print(id, bill)
	
	if check_id(id) == 'no':
		return 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select set_bill(%s, %s)", (int(id), float(bill)))
	row = cursor.fetchone()
	if row[0] == 'ok':
		conn.commit()
	cursor.close()
	return row[0]

def add_extra(id, extra):
	print('add_extra')
	print(id, extra)

	if check_id(id) == 'no':
		return 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select add_extra(%s, %s)", (int(id), float(extra)))
	row = cursor.fetchone()
	if row[0] == 'ok':
		conn.commit()
	cursor.close()
	return row[0]


def check_client(client):
	print('check_client')
	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select check_client(%s)", (client, ))
	row = cursor.fetchone()
	cursor.close()
	return row[0]

def check_person(client, person):

	print('check_person')
	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select check_person(%s, %s)", (client, person))
	row = cursor.fetchone()
	cursor.close()
	return row[0]

def add_client(client_data):
	print('add_client')
	print(client_data)
	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select add_client(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", dict_to_list(client_data))
	conn.commit()
	row = cursor.fetchone()
	cursor.close()
	print(row[0])
	return row[0]


def update_client(client_data):
	print('update_client')
	print(client_data)
	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select update_client(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", dict_to_list(client_data))
	conn.commit()
	row = cursor.fetchone()
	cursor.close()
	print(row[0])
	return row[0]

def add_person(person_data):
	print('add_person')
	print(person_data)
	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select add_person(%s, %s, %s, %s, %s, %s)", dict_to_list(person_data))
	conn.commit()
	row = cursor.fetchone()
	cursor.close()
	print(row[0])
	return row[0]

def update_person(person_data):
	print('update_person')
	print(person_data)
	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select update_person(%s, %s, %s, %s, %s, %s)", dict_to_list(person_data))
	conn.commit()
	row = cursor.fetchone()
	cursor.close()
	print(row[0])
	return row[0]

def get_pay_info(id):
	
	pay_info = {}

	if check_id(id) == 'no':
		return pay_info, 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select * from ord_money where id = %s", (id,))
	pay_info = cursor.fetchone()
	
	if pay_info is None:
		cursor.close()
		return pay_info, 'no'

	cursor.close()
	return pay_info, 'ok'

def get_ord_content(id):
	
	ord_content = {}

	if check_id(id) == 'no':
		return ord_content, 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select * from ord_content where id = %s", (id,))
	ord_content = cursor.fetchone()
	
	if ord_content is None:
		cursor.close()
		return ord_content, 'no'

	cursor.close()
	return ord_content, 'ok'

def get_ord_history(id):
	
	ord_history = {}

	if check_id(id) == 'no':
		return ord_history, 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select * from orders where id = %s", (id,))
	ord_history = cursor.fetchone()
	
	if ord_history is None:
		cursor.close()
		return ord_history, 'no'

	cursor.close()
	return ord_history, 'ok'

def get_ord_client(id):
	
	ord_client = {}

	if check_id(id) == 'no':
		return ord_client, 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select client, person from orders where id = %s", (id,))
	cp = cursor.fetchone()
	
	if cp is None:
		cursor.close()
		return ord_client, 'no'

	cursor.execute("select tel from persons where client = %s and name = %s", (cp['client'], cp['person']))
	tel = cursor.fetchone()

	if tel is None:
		cursor.execute("select tel from clients where name = %s", (cp['client'], ))
		tel = cursor.fetchone()

	ord_client['client'] = cp['client']
	ord_client['person'] = cp['person']
	ord_client['tel'] = tel['tel']

	cursor.close()
	return ord_client, 'ok'

def get_client(name):

	client = {}

	if check_client(name) == 'no':
		return client, 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select * from clients where name = %s", (name,))
	client = cursor.fetchone()
	
	if client is None:
		cursor.close()
		return client, 'no'

	cursor.close()
	return client, 'ok'

def get_person(client, name):

	person = {}

	if check_person(client, name) == 'no':
		return person, 'no'

	cursor = conn.cursor(cursor_factory=DictCursor)
	cursor.execute("select * from persons where name = %s and client = %s", (name, client))
	person = cursor.fetchone()
	
	if person is None:
		cursor.close()
		return person, 'no'

	cursor.close()
	return person, 'ok'



##########################################	GUI WINDOWS	#############################################
sg.theme('Light Purple')

def main_window():
	last_id = get_last_id()
	layout = [[sg.Text('Номер последнего заказа: '), sg.Text(last_id, key = '-LAST_ID-', size = (10, 1), font = 'Any 15'), sg.Button('Обновить номер')], 
	[sg.Text('      ')],
	[sg.Text('Добавить заказ: ', size = (30, 1)), sg.Input(key = '-ORDER-',  size = (50, 1)), sg.FileBrowse(), sg.Text('	'), sg.Button('Оплачен')],
	[sg.Text('		')],
	[sg.Text('      ')],
	[sg.Text('Номер заказа: ', size = (30, 1)), sg.Input(key = '-ID-', size = (50, 1))], 
	[sg.Text('		')],
	[sg.Button('Узнать статус', size = (27, 1)), sg.Button('Изменить статус', size = (27, 1))],
	[sg.Text('Статус: ', size = (30, 1)), sg.Input(key = '-SHOW_STATUS-', size = (50, 1))], 
	[sg.Text('		')],
	[sg.Button('Оплата', size = (27, 1))],
	[sg.Text('Цена клиенту: ', size = (30, 1)), sg.Input(key = '-PRICE-', size = (50, 1))],
	[sg.Text('Оплачено: ', size = (30, 1)), sg.Input(key = '-PREPAY-', size = (50, 1)), sg.Button('Добавить к оплате', size = (27, 1)), sg.Input(key = '-ADD_PREPAY-', size = (20, 1))],
	[sg.Text('Долг: ', size = (30, 1)), sg.Input(key = '-DEBT-', size = (50, 1))],
	[sg.Text('Способ оплаты: ', size = (30, 1)), sg.Input(key = '-PAYDOC-', size = (50, 1))],
	[sg.Text('Счет: ', size = (30, 1)), sg.Input(key = '-BILL-', size = (50, 1)), sg.Button('Выставить сумму', size = (27, 1)), sg.Input(key = '-SET_BILL-', size = (20, 1))],
	[sg.Text('Доставка: ', size = (30, 1)), sg.Input(key = '-DELIVERY-', size = (50, 1))],
	[sg.Text('Непредвиденные: ', size = (30, 1)), sg.Input(key = '-EXTRA-', size = (50, 1)), sg.Button('Добавить расходы', size = (27, 1)), sg.Input(key = '-ADD_EXTRA-', size = (20, 1))],
	[sg.Text('Прибыль: ', size = (30, 1)), sg.Input(key = '-PROFIT-', size = (50, 1))],
	[sg.Text('		')],
	[sg.Button('Состав заказа', size = (27, 1)), sg.Text(' ', size = (50, 1)), sg.Button('История', size = (27, 1)), sg.Text(' ', size = (50, 1))],
	[sg.Text('Цвет: ', size = (30, 1)), sg.Input(key = '-COLOR-', size = (50, 1)), sg.Text('Оплачен: ', size = (30, 1)), sg.Input(key = '-PAY_DATE-', size = (50, 1))],
	[sg.Text('Рисунок: ', size = (30, 1)), sg.Input(key = '-SHAPE-', size = (50, 1)), sg.Text('Заявка: ', size = (30, 1)), sg.Input(key = '-REQUEST_DATE-', size = (50, 1))],
	[sg.Text('Торец: ', size = (30, 1)), sg.Input(key = '-EDGE-', size = (50, 1)), sg.Text('Получен: ', size = (30, 1)), sg.Input(key = '-RECIEVE_DATE-', size = (50, 1))],
	[sg.Text('М2: ', size = (30, 1)), sg.Input(key = '-MSQ-', size = (50, 1)), sg.Text('Выдать к: ', size = (30, 1)), sg.Input(key = '-GIVE_DATE-', size = (50, 1))],
	[sg.Text('Кол-во штук: ', size = (30, 1)), sg.Input(key = '-ITEMS-', size = (50, 1)), sg.Text('', size = (30, 1))],
	[sg.Text('		')],
	[sg.Button('Заказчик', size = (27, 1)), sg.Text(' ', size = (50, 1)), sg.Button('Добавить клиента', size = (27, 1)), sg.Button('Добавить представителя', size = (47, 1))],
	[sg.Text('Организация: ', size = (30, 1)), sg.Input(key = '-CLIENT-', size = (50, 1)), sg.Button('Найти клиента', size = (27, 1))],
	[sg.Text('Контактное лицо: ', size = (30, 1)), sg.Input(key = '-PERSON-', size = (50, 1))],
	[sg.Text('Телефон: ', size = (30, 1)), sg.Input(key = '-TEL-', size = (50, 1))],
	[sg.Text('		')],
	[sg.Button('Выход')]]

	return sg.Window('Фасады', layout)

def pre_pay_data_window():
	layout = [[sg.Text('Платежный документ:', size = (40, 1)), sg.InputCombo(('Счет', 'Приходный ордер'), key = '-PAYDOC_TYPE-', size = (29, 1))], 
	[sg.Text('Документ: ', size = (40, 1)), sg.Input(key = '-PAYDOC-', size = (30, 1))], 
	[sg.Text('Дата документа (yyyy-mm-dd):', size = (40, 1)), sg.Input(key = '-PAYDOC_DATE-', size = (30, 1))],
	[sg.Text('		')],
	[sg.Button('Добавить'), sg.Button('Отмена')]]

	return sg.Window('Оплата', layout)

def full_pay_data_window(id):
	layout = [[sg.Text('Номер заказа: ', size = (40, 1)), sg.Text(id, size = (30, 1))], 
	[sg.Text('Платежный документ:', size = (40, 1)), sg.InputCombo(('Счет', 'Приходный ордер'), key = '-PAYDOC_TYPE-', size = (29, 1))], 
	[sg.Text('Документ: ', size = (40, 1)), sg.Input(key = '-PAYDOC-', size = (30, 1))], 
	[sg.Text('Дата документа (yyyy-mm-dd):', size = (40, 1)), sg.Input(key = '-PAYDOC_DATE-', size = (30, 1))],
	[sg.Text('Внесена оплата, руб.', size = (40, 1)), sg.Input(key = '-PREPAY-', size = (30, 1))],
	[sg.Text('		')],
	[sg.Button('Добавить')]]

	return sg.Window('Оплата', layout)

def add_client_window(client = '', person = '', tel = ''):
	layout = [[sg.Text('Организация: ', size = (50, 1)), sg.Input(key = '-NAME-', default_text = client, size = (40, 1))], 
	[sg.Text('Категория: ', size = (50, 1)), sg.InputCombo(('Компания', 'Частный', 'Дизайнер'), key = '-CATEGORY-', size = (39, 1))],
	[sg.Text('Ярлык: ', size = (50, 1)), sg.InputCombo(('Новый', 'Постоянный', 'Лояльный', 'Отказной', 'Привлекаем'), key = '-TAG-', size = (39, 1))],
	[sg.Text('Город: ', size = (50, 1)), sg.Input(key = '-CITY-', size = (40, 1))],
	[sg.Text('Адрес: ', size = (50, 1)), sg.Input(key = '-ADDRESS-', size = (40, 1))], 
	[sg.Text('Телефон: ', size = (50, 1)), sg.Input(key = '-TEL-', default_text = tel, size = (40, 1))],
	[sg.Text('Email: ', size = (50, 1)), sg.Input(key = '-MAIL-', size = (40, 1))],
	[sg.Text('Сайт: ', size = (50, 1)), sg.Input(key = '-SITE-', size = (40, 1))],
	[sg.Text('Instagram: ', size = (50, 1)), sg.Input(key = '-INST-', size = (40, 1))],
	[sg.Text('ВК: ', size = (50, 1)), sg.Input(key = '-VK-', size = (40, 1))],
	[sg.Text('Последнее взаимодействие: ', size = (50, 1)), sg.Input(size = (40, 1), key = '-LAST_CONTACT-', default_text = str(datetime.datetime.now().date()))],
	[sg.Text('		')],
	[sg.Button('Добавить'), sg.Button('Отмена')], 
	[sg.Text('		')],
	[sg.Button('Добавить представителя')]]

	return sg.Window('Новый клиент', layout)

def add_person_window(client = '', person = '', tel = ''):
	layout = [[sg.Text('Организация: ', size = (50, 1)), sg.Input(key = '-CLIENT-', default_text = client, size = (50, 1))],
	[sg.Text('Имя: ', size = (50, 1)), sg.Input(key = '-NAME-', default_text = person, size = (50, 1))],
	[sg.Text('Телефон: ', size = (50, 1)), sg.Input(key = '-TEL-', default_text = tel, size = (50, 1))], 
	[sg.Text('Email: ', size = (50, 1)), sg.Input(key = '-MAIL-', size = (50, 1))],
	[sg.Text('Последнее взаимодействие: ', size = (50, 1)), sg.Input(size = (50, 1), key = '-LAST_CONTACT-', default_text = str(datetime.datetime.now().date()))],
	[sg.Text('Дополнительно: ', size = (50, 1)), sg.Input(key = '-INFO-', size = (50, 1))],
	[sg.Text('		')],
	[sg.Button('Добавить'), sg.Button('Отмена')]]

	return sg.Window('Новый представитель', layout) 

def status_update_window(current_status, current_date, id):
	statuses = ('Оплачен', 'Заказан', 'На складе', 'Выдан')

	new_status = ''
	if current_status == 'Оплачен':
		new_status = 'Заказан'
	if current_status == 'Заказан':
		new_status = 'На складе'
	if current_status == 'На складе':
		new_status = 'Выдан'

	new_date = str(datetime.datetime.now().date())

	layout = [[sg.Text('Номер заказа: ', size = (30, 1)), sg.Text(id)],
	[sg.Text('Текущий статус: ', size = (30, 1)), sg.Text(current_status)],
	[sg.Text('Дата: ', size = (30, 1)), sg.Text(current_date)], 
	[sg.Text('Новый статус: ', size = (30, 1)), sg.InputCombo(values = statuses, key = '-NEW_STATUS-', default_value = new_status, size = (29, 1))],
	[sg.Text('Дата статуса: ', size = (30, 1)), sg.Input(key = '-NEW_DATE-', default_text = new_date, size = (30, 1))],
	[sg.Button('Обновить'), sg.Button('Отмена')]]

	return sg.Window('Изменить статус заказа', layout)


def find_client_window():
	layout = [[sg.Text('Организация: ', size = (50, 1)), sg.Input(key = '-NAME-', default_text = "Введите название", size = (40, 1))], 
	[sg.Text('Категория: ', size = (50, 1)), sg.Input(key = '-CATEGORY-', size = (40, 1))],
	[sg.Text('Ярлык: ', size = (50, 1)), sg.Input(key = '-TAG-', size = (40, 1))],
	[sg.Text('Город: ', size = (50, 1)), sg.Input(key = '-CITY-', size = (40, 1))],
	[sg.Text('Адрес: ', size = (50, 1)), sg.Input(key = '-ADDRESS-', size = (40, 1))], 
	[sg.Text('Телефон: ', size = (50, 1)), sg.Input(key = '-TEL-', size = (40, 1))],
	[sg.Text('Email: ', size = (50, 1)), sg.Input(key = '-MAIL-', size = (40, 1))],
	[sg.Text('Сайт: ', size = (50, 1)), sg.Input(key = '-SITE-', size = (40, 1))],
	[sg.Text('Instagram: ', size = (50, 1)), sg.Input(key = '-INST-', size = (40, 1))],
	[sg.Text('ВК: ', size = (50, 1)), sg.Input(key = '-VK-', size = (40, 1))],
	[sg.Text('Последнее взаимодействие: ', size = (50, 1)), sg.Input(size = (40, 1), key = '-LAST_CONTACT-')],
	[sg.Text('		')],
	[sg.Button('Найти')], 
	[sg.Text('		')],
	[sg.Button('Найти представителя'), sg.Input(key = '-PERSON-', size = (40, 1), default_text = "Введите имя")],
	[sg.Text('Телефон: ', size = (50, 1)), sg.Input(key = '-PERSON_TEL-', size = (50, 1))], 
	[sg.Text('Email: ', size = (50, 1)), sg.Input(key = '-PERSON_MAIL-', size = (50, 1))],
	[sg.Text('Последнее взаимодействие: ', size = (50, 1)), sg.Input(size = (50, 1), key = '-PERSON_LAST_CONTACT-')],
	[sg.Text('Дополнительно: ', size = (50, 1)), sg.Input(key = '-INFO-', size = (50, 1))]]

	return sg.Window('Найти клиента', layout)


########################################## APP DATA GETTING FUNCTIONS	##############################

## data from window_pd 
# def get_pre_pay_data(value):
# 	pay_data = {}
# 	pay_data['type'] = value['-PAYDOC_TYPE-']
# 	pay_data['doc'] = value['-PAYDOC-']
# 	pay_data['date'] = value['-PAYDOC_DATE-']

#	return pay_data

def get_full_pay_data(value):
	pay_data = {}
	pay_data['type'] = value['-PAYDOC_TYPE-']
	pay_data['doc'] = value['-PAYDOC-']
	pay_data['date'] = value['-PAYDOC_DATE-']
	pay_data['prepay'] = value['-PREPAY-']

	return pay_data

def get_client_data(value):
	client_data = {}
	client_data['name'] = value['-NAME-']
	client_data['category'] = value['-CATEGORY-']
	client_data['tag'] = value['-TAG-']
	client_data['city'] = value['-CITY-']
	client_data['address'] = value['-ADDRESS-']
	client_data['tel'] = value['-TEL-']
	client_data['mail'] = value['-MAIL-']
	client_data['site'] = value['-SITE-']
	client_data['inst'] = value['-INST-']
	client_data['vk'] = value['-VK-']
	client_data['last_contact'] = value['-LAST_CONTACT-']

	return client_data

def get_person_data(value):
	person_data = {}
	person_data['name'] = value['-NAME-']
	person_data['client'] = value['-CLIENT-']
	person_data['tel'] = value['-TEL-']
	person_data['mail'] = value['-MAIL-']
	person_data['info'] = value['-INFO-']
	person_data['last_contact'] = value['-LAST_CONTACT-']

	return person_data


## data from excel order form
def read_order_xls(order_file):
	d = {}

	wb = xlrd.open_workbook(order_file)
	sheet = wb.sheet_by_index(0) 

	#cell_value(row, col)
	d['client'] = sheet.cell_value(3, 3) #заказчик
	d['person'] = sheet.cell_value(3, 5) #имя контактного лица
	d['tel'] = str(int(sheet.cell_value(3, 6))) #телефон коетактного лица
	d['id'] = int(sheet.cell_value(4, 3))

	date = sheet.cell_value(4, 5)
	date_dt = datetime.datetime(*xlrd.xldate_as_tuple(date, wb.datemode))
	d['date'] = str(date_dt.date())

	d['shape'] = sheet.cell_value(7, 3) #рисунок
	d['edge'] = sheet.cell_value(8, 3) #форма торца
	d['color'] = sheet.cell_value(9, 3) #цвет

	d['type'] = 'Пленка' #пленка или крашеный
	if ('RAL ' in d['color']) or ('NCS ' in d['color']):
		d['type'] = 'Краска'

	d['items'] = int(sheet.cell_value(52, 5)) #штук в заказе
	d['msq'] = round(sheet.cell_value(52, 6), 3) #площадь м2

	price = sheet.cell_value(52, 7) #цена
	dr, c = math.modf(price)
	d['price'] = int(c)
	if dr >= 0.5: 
		d['price'] = int(c + 1)

	return d

########################################## MAIN LOOP ######################################

window = None
#connection to DB
print("connecting to DB...")
conn = pg.connect(dbname = 'firstdb', user = 'yanadb', password = '171007', host = 'localhost')

while True:
	if window is None:
		window = main_window()
	event, value = window.read()

	if event in (None, 'Выход'):
		conn.close()
		break

	if event == 'Обновить номер':
		pass
		print('Обновить номер')
		last_id = get_last_id()
		window['-LAST_ID-'].update(last_id)

	# if event == 'Предварительно':
	# 	if value['-ORDER-'] == '':
	# 		sg.popup_error('Файл заказа не выбран')
	# 	else:
	# 		print('Предварительно')
	# 		#label = sg.popup('Метка заказа')
	# 		print(value['-ORDER-'])
	# 		order_data = read_order_xls(value['-ORDER-'])
	# 		print(order_data)
	# 		label = sg.popup_get_text('Название заказа: ')
	# 		if label == '':
	# 			label = sg.popup_get_text('Необходимо название заказа: ')
	# 		print(label)
	# 		pay_data = None
	# 		answ = sg.popup_yes_no('Добавить данные об оплате?')
	# 		if answ == 'Yes':
	# 			window_ppd = pre_pay_data_window()
	# 			event_ppd, value_ppd = window_ppd.read()

	# 			if event_ppd == 'Добавить':
	# 				pay_data = get_pre_pay_data(value_ppd)
	# 				sg.popup('Данные об оплате добавлены')
	# 				window_ppd.close()
	# 				window_ppd = None

	# 			if event_ppd == 'Отмена': 
	# 				sg.popup('Заказ без данных об оплате')
	# 				window_ppd.close()
	# 				window_ppd = None

	# 			if event_ppd is None: 
	# 				window_ppd.close()
	# 				window_ppd = None

	# 		if answ in (None, 'No'):
	# 			sg.popup('Заказ без данных об оплате')

	# 		if add_pre_order(order_data, pay_data, label) == 'ok':
	# 			sg.popup('Предварительный заказ добавлен')
	# 		else:
	# 			sg.popup_error('Заказ не добавлен')

	if event == 'Оплачен':
		if value['-ORDER-'] == '':
			sg.popup_error('Файл заказа не выбран')
		else:
			print('Оплачен')
			print(value['-ORDER-'])
			order_data = read_order_xls(value['-ORDER-'])
			print(order_data)

			client = order_data['client']
			person = order_data['person']
			tel = order_data['tel']

			no_client = 0

			wr_id = 0

			if(check_id(order_data['id']) == 'ok'):
				sg.popup_error('Заказ с номером ' + str(order_data['id']) + ' уже существует')
				wr_id = 1


			if(check_client(client) == 'no'):
				sg.popup_error('Добавьте клиента ' + client + ' в базу данных')
				no_client = 1


			if(check_person(client, person) == 'no'):
				sg.popup_error('Добавьте представителя ' + person + ' компании ' + client + ' в базу данных')
				no_client = 1

			if((no_client == 1) or (wr_id == 1)):
				sg.popup_error('Заказ не добавлен')
			else:

				pay_data = None
				window_pd = full_pay_data_window(order_data['id'])
				event_pd, value_pd = window_pd.read()

				if event_pd == 'Добавить':
					pay_data = get_full_pay_data(value_pd)
					sg.popup('Данные об оплате добавлены')
					window_pd.close()
					window_pd = None

				if event_pd == 'Отмена': 
					sg.popup_error('Заказ не добавлен')
					window_pd.close()
					window_pd = None

				if event_pd is None: 
					sg.popup_error('Заказ не добавлен')
					window_pd.close()
					window_pd = None

				if pay_data is None:
					sg.popup_error('Заказ не добавлен: нет данных об оплате')
				else:
					if(add_order(order_data, pay_data) == 'ok'):
						order_str = 'Заказ номер ' + str(order_data['id']) + ' добавлен'
						sg.popup(order_str)
					else:
						order_str = 'Заказ номер ' + str(order_data['id']) + ' не добавлен'
						sg.popup_error(order_str)

	if event == 'Узнать статус':
		id = value['-ID-']
		if id == '':
			sg.popup_error('Введите номер заказа')
		else:
			status, date = get_status_with_date(id)
			if status == 'no':
				sg.popup_error('Заказ не найден!')
			else:
				window['-SHOW_STATUS-'].update(status + '  '+ date)

	if event == 'Изменить статус':
		id = value['-ID-']
		if id == '':
			sg.popup_error('Введите номер заказа')
		else:
			status, date = get_status_with_date(id)
			if status == 'no':
				sg.popup_error('Заказ не найден')
			else:
				window_us = status_update_window(status, date, id)
				event_us, value_us = window_us.read()

				if event_us in (None, 'Отмена'):
					window_us.close()
					window_us = None

				if  event_us == 'Обновить':
					new_status = value_us['-NEW_STATUS-']
					new_date = value_us['-NEW_DATE-']
					if (new_status not in ('Оплачен', 'Заказан', 'На складе', 'Выдан')) or (new_date == ''):
						sg.popup_error('Не корректные значения')
					else:
						if (update_status(new_status, new_date, id) == 'ok'):
							sg.popup('Статус заказа номеер ' + id + ' обновлен')
							window_us.close()
							window_us = None
						else:
							sg.popup_error('Не удалось обновить статус')
							window_us.close()
							window_us = None

	if event == 'Оплата':
		id = value['-ID-']
		if id == '':
			sg.popup_error('Введите номер заказа')
		else:
			pay_info, res = get_pay_info(id)
			if res != 'ok':
				sg.popup_error('Заказ не найден')
			else: 
				window['-PRICE-'].update(pay_info['price'])
				window['-PREPAY-'].update(pay_info['prepay'])
				window['-DEBT-'].update(pay_info['debt'])
				window['-PAYDOC-'].update(pay_info['doc'])
				window['-BILL-'].update(pay_info['bill'])
				window['-DELIVERY-'].update(pay_info['delivery'])
				window['-EXTRA-'].update(pay_info['extra'])
				window['-PROFIT-'].update(pay_info['profit'])

	if event == 'Добавить к оплате':
		id = value['-ID-']
		if id == '':
			sg.popup_error('Введите номер заказа')
		else:
			prepay = value['-ADD_PREPAY-']
			if prepay == '':
				sg.popup_error('Введите сумму')
			else:
				if(add_prepay(id, prepay) == 'ok'):
					sg.popup('Сумма ' + prepay + ' добавлена к оплате заказа номер ' + id)
				else: 
					sg.popup_error('Заказ не найден')

	if event == 'Выставить сумму':
		id = value['-ID-']
		if id == '':
			sg.popup_error('Введите номер заказа')
		else:
			bill = value['-SET_BILL-']
			if bill == '':
				sg.popup_error('Введите сумму')
			else:
				if(set_bill(id, bill) == 'ok'):
					sg.popup('Выставлена сумма счета ' + bill + ' для заказа номер ' + id)
				else: 
					sg.popup_error('Заказ не найден')


	if event == 'Добавить расходы':
		id = value['-ID-']
		if id == '':
			sg.popup_error('Введите номер заказа')
		else:
			extra = value['-ADD_EXTRA-']
			if extra == '':
				sg.popup_error('Введите сумму')
			else:
				if(add_extra(id, extra) == 'ok'):
					sg.popup('Добавлены расходы ' + extra + ' для заказа номер ' + id)
				else: 
					sg.popup_error('Заказ не найден')

	if event == 'Состав заказа':
		id = value['-ID-']
		if id == '':
			sg.popup_error('Введите номер заказа')
		else:
			ord_content, res = get_ord_content(id)
			if res != 'ok':
				sg.popup_error('Заказ не найден')
			else: 
				window['-COLOR-'].update(ord_content['color'])
				window['-SHAPE-'].update(ord_content['shape'])
				window['-EDGE-'].update(ord_content['edge'])
				window['-MSQ-'].update(ord_content['msq'])
				window['-ITEMS-'].update(ord_content['items'])

	if event == 'История':
		id = value['-ID-']
		if id == '':
			sg.popup_error('Введите номер заказа')
		else:
			ord_history, res = get_ord_history(id)
			if res != 'ok':
				sg.popup_error('Заказ не найден')
			else: 
				window['-PAY_DATE-'].update(str(ord_history['pay_date']))
				window['-REQUEST_DATE-'].update(str(ord_history['request_date']))
				window['-RECIEVE_DATE-'].update(str(ord_history['recieve_date']))
				window['-GIVE_DATE-'].update(str(ord_history['give_date']))

	if event == 'Заказчик':
		id = value['-ID-']
		if id == '':
			sg.popup_error('Введите номер заказа')
		else:
			ord_client, res = get_ord_client(id)
			if res != 'ok':
				sg.popup_error('Заказ не найден')
			else: 
				window['-CLIENT-'].update(ord_client['client'])
				window['-PERSON-'].update(ord_client['person'])
				window['-TEL-'].update(ord_client['tel'])

	if event == 'Добавить клиента':
		window_ac = add_client_window()
		event_ac, value_ac = window_ac.read()

		if event_ac in (None, 'Отмена'):
			window_ac.close()
			window_ac = None

		if event_ac == 'Добавить':
			client_data = get_client_data(value_ac)
			if(check_client(client_data['name']) == 'no'):
				if(add_client(client_data) == 'ok'):
					sg.popup('Клиент ' + client_data['name'] + ' добавлен')
				else:
					sg.popup_error('Клиент ' + client_data['name'] + ' не добавлен')

				window_ac.close()
				window_ac = None

			else:
				answ = sg.popup_yes_no('Клиент ' + client_data['name'] + ' есть в базе. Обновить данные?')
				if answ == 'Yes':
					if(update_client(client_data) == 'ok'):
						sg.popup('Данные ' + client_data['name'] + ' обновлены')
					else:
						sg.popup_error('Данные ' + client_data['name'] + ' не обновлены')

				window_ac.close()
				window_ac = None

	if event == 'Добавить представителя':
		window_ac = add_person_window()
		event_ac, value_ac = window_ac.read()

		if event_ac in (None, 'Отмена'):
			window_ac.close()
			window_ac = None

		if event_ac == 'Добавить':
			person_data = get_person_data(value_ac)

			if(check_client(person_data['client']) == 'yes'):
				sg.popup_error('Организации ' + person_data['client'] + ' нет в базе')
				window_ac.close()				
				window_ac = None
			else:

				if(check_person(person_data['client'], person_data['name']) == 'no'):
					if(add_person(person_data) == 'ok'):
						sg.popup('Представитель ' + person_data['name'] + ' организации ' + person_data['client'] + ' добавлен')
					else:
						sg.popup_error('Представитель ' + person_data['name'] + ' организации ' + person_data['client'] + ' не добавлен')

					window_ac.close()
					window_ac = None
					
				else:
					answ = sg.popup_yes_no('Представитель ' + person_data['name'] + ' организации ' + person_data['client'] + ' есть в базе. Обновить данные?')
					if answ == 'Yes':
						if(update_person(person_data) == 'ok'):
							sg.popup('Данные ' + person_data['name'] + ' организации ' + person_data['client'] + ' обновлены')
						else:
							sg.popup_error('Данные ' + person_data['name'] + ' организации ' + person_data['client'] + ' не обновлены')

					window_ac.close()
					window_ac = None

	if event == 'Найти клиента':
		window_fc = None
		while True:
			if window_fc is None:
				window_fc = find_client_window()
			event_fc, value_fc = window_fc.read()

			if event_fc is None:
				window_fc.close()
				window_fc = None
				break

			if event_fc == 'Найти':
				if value_fc['-NAME-'] in ('', 'Введите название'):
					sg.popup_error("Введите название")
				else:
					client, res = get_client(value_fc['-NAME-'])
					if(res == 'no'):
						sg.popup_error("Клиент не найден")
					else:
						window_fc['-CATEGORY-'].update(client['category'])
						window_fc['-TAG-'].update(client['tag'])
						window_fc['-CITY-'].update(client['city'])
						window_fc['-ADDRESS-'].update(client['address'])
						window_fc['-TEL-'].update(client['tel'])
						window_fc['-MAIL-'].update(client['mail'])
						window_fc['-SITE-'].update(client['site'])
						window_fc['-INST-'].update(client['inst'])
						window_fc['-VK-'].update(client['vk'])
						window_fc['-LAST_CONTACT-'].update(client['last_contact'])
					#update field

			if event_fc == 'Найти представителя':
				if value_fc['-NAME-'] in ('', 'Введите название'):
					sg.popup_error("Введите название организации")

				else:
					if value_fc['-PERSON-'] in ('', 'Введите имя'):
						sg.popup_error("Введите имя")

					else:
						person, res = get_person(value_fc['-NAME-'], value_fc['-PERSON-'])
						if(res == 'no'):
							sg.popup_error("Представитель не найден")
						else:
							window_fc['-PERSON_TEL-'].update(person['tel'])
							window_fc['-PERSON_MAIL-'].update(person['mail'])
							window_fc['-PERSON_LAST_CONTACT-'].update(person['last_contact'])
							window_fc['-INFO-'].update(person['info'])






