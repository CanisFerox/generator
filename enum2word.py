#! /usr/bin/python3

import re
import argparse
import uno
import os


def create_parser():
	parser = argparse.ArgumentParser(
		prog="enum",
		description="Программа выполняет поиск криминалистически значимой информации в HKLM\\SYSTEM\\ControlSet*\\Enum\\ и вывод полученных данных в виде html документа.",
		epilog="Created by Canis (canis.ferox@yandex.ru)"
		)
	parser.add_argument("-f", "--file", required=True, help="Путь к экспортированному файлу ветви реестра HKLM")
	parser.add_argument("-s", "--section", required=True, choices=["USB", "USBSTOR", "SCSI"] ,help="Название ветви реестра для анализа")
	parser.add_argument("-o", "--output", help="Имя файла для сохранения отчета (без расширения)")
	return parser

def get_devices(f, section):
	item_re = re.compile(r"SYSTEM\\ControlSet[0-9]{3}\\Enum\\" + section + r"\\([^\\^\s]+)\\([^\\^\s]+)\"")
	not_item_re = re.compile(r"\\SYSTEM\\ControlSet\d*\\Enum\\([^\\^\s]+)\\([^\\^\s]+)\\([^\\^\s]+)")
	time_re = re.compile("(\d+\-\d+\-\d+\s*\S\s*\d+:\d+:\d+)")
	params_re = {
		"DeviceDesc": re.compile("\"DeviceDesc\""),
		"LocationInformation": re.compile("\"LocationInformation\""),
		"FriendlyName": re.compile("\"FriendlyName\"")
	}
	param_value_re = re.compile("\S+\s+\"(.+)\"$")
	devices = {}
	vp_id, instance_id = None, None
	for line in f:
		item = item_re.search(line)
		not_item = not_item_re.search(line)
		if item is not None:
			vp_id, instance_id = item.groups()
			if devices.get(vp_id) is None:
				devices[vp_id] = {}
			if devices[vp_id].get(instance_id) is None:
				devices[vp_id][instance_id] = {}
			new_str = f.readline()
			devices[vp_id][instance_id]["time"] = time_re.search(new_str).group(1)	# сохранение временной метки
			continue
		elif item is None and not_item is not None:
			vp_id = None
			continue
		for key_re in params_re.keys():
			item = params_re[key_re].search(line)
			if item is not None and vp_id is not None:
				new_str = f.readline()	# пропускаем строку с типом параметра
				new_str = f.readline()
				new_str = f.readline()
				temp_res = param_value_re.search(new_str)
				if devices[vp_id][instance_id].get(key_re) is None:
					devices[vp_id][instance_id][key_re] = temp_res.group(1) if temp_res is not None else "-"
	return devices

def insertTextIntoCell(table, cellName, text):
	tableText = table.getCellByName(cellName)
	cursor = tableText.createTextCursor()
	tableText.setString(text)

def generate_table(items, namespace):
	local = uno.getComponentContext()
	resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
	context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
	desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
	url = "private:factory/swriter"
	document = desktop.loadComponentFromURL(url, "_blank", 0, ())
	cursor = document.Text.createTextCursor()

	################# DOC CONTENT #################

	text = "Данные извлечены из раздела SYSTEM\\ControlSet\\Enum\\" + namespace.section + ".\n\n"       # Table Header
	document.Text.insertString(cursor, text, 0)

	key_count = 1
	for key in items.keys():
		for subkey in items.get(key).keys():
			key_count += 1

	table = document.createInstance("com.sun.star.text.TextTable")
	table.initialize(key_count, 5)
	document.Text.insertTextContent(cursor, table, 0)

	insertTextIntoCell(table, "A1", "№")
	insertTextIntoCell(table, "B1", "Id производителя и продукта")
	insertTextIntoCell(table, "C1", "Id устройства")
	insertTextIntoCell(table, "D1", "Информация о подключении")
	insertTextIntoCell(table, "E1", "Дата последнего подключения")

	column = ["A", "B", "C", "D", "E", "F", "G"]
	row = 2

	for key in items.keys():
		for subkey in items.get(key).keys():
			insertTextIntoCell(table, column[0] + str(row), str(row - 1))
			insertTextIntoCell(table, column[1] + str(row), key)
			insertTextIntoCell(table, column[2] + str(row), subkey)
			value = items.get(key).get(subkey).get("LocationInformation") if items.get(key).get(subkey).get("LocationInformation") is not None else "-"
			insertTextIntoCell(table, column[3] + str(row), value)
			insertTextIntoCell(table, column[4] + str(row), items.get(key).get(subkey).get("time"))
			row += 1

	###############################################

	if namespace.output is not None:
		document.storeAsURL("file:///" + os.getcwd() + "/" + namespace.output + ".odt", ())
		document.dispose()

def main(namespace):
	with open(namespace.file, "r", encoding="UTF-8") as f:
		items = get_devices(f, namespace.section)
		generate_table(items, namespace)
	if namespace.output is not None:
		print("Готово. Файл отчета сохранен с именем \"{}/{}.odt\"".format(os.getcwd(), namespace.output))
	else:
		print("Готово. Файл отчета сформирован в новом документе LibreOffice, вам необходимо сохранить его вручную.")

if __name__ == "__main__":
	parser = create_parser()
	namespace = parser.parse_args()
	main(namespace)
