#!/usr/bin/python3

import sys
import argparse
import re
import os

def create_parser():
	parser = argparse.ArgumentParser(
		prog="enum",
		description="Программа выполняет поиск криминалистически значимой информации в HKLM\\SYSTEM\\ControlSet001\\Enum\\ и вывод полученных данных в виде html документа.",
		epilog="Created by Canis (canis.ferox@yandex.ru)"
		)
	parser.add_argument("-f", "--file", required=True, help="Путь к экспортированному файлу ветви реестра HKLM")
	parser.add_argument("-s", "--section", required=True, choices=["USB", "USBSTOR", "SCSI"] ,help="Название ветви реестра для анализа")
	parser.add_argument("-o", "--output", required=True, help="Имя файла для сохранения отчета (без расширения)")
	return parser


def main(namespace):
	with open(namespace.file, "r", encoding="UTF-16") as f:
		devices = get_devices(f, namespace.section)
		generate_report(devices, namespace)
	print("Готово. Файл отчета сохранен с именем \"{}/{}.html\"".format(os.getcwd(), namespace.output))


def generate_report(devices, namespace):
	html  = "<head><title> Результат анализа ветви HKLM\\SYSTEM\\ControlSet001\\Enum\\" + namespace.section + "</title></head>\n"
	html += "<body>\n <table border=\"1\" cellpadding=\"2\" cellspacing=\"0\">\n"
	html += "<tr> <th>№</th><th>Id производителя и продукта</th><th>Id устройства</th><th>Название устройства</th><th>Описание</th><th>Инф о подключении</th><th>Дата последнего подключения</th> </tr>\n"
	n = 1
	for vpkey in devices.keys():
		for ikey in devices[vpkey].keys():
			html += "<tr> <td>{}</td><td>{}</td>".format(n, vpkey)
			n += 1
			html += "<td>{}</td>".format(ikey)
			temp = devices[vpkey][ikey]
			html += "<td>{}</td><td>{}</td><td>{}</td><td>{}</td> </tr>\n".format(
				temp.get("FriendlyName") if temp.get("FriendlyName") != None else "-", 
				temp.get("DeviceDesc") if temp.get("DeviceDesc") != None else "-", 
				temp.get("LocationInformation") if temp.get("LocationInformation") != None else "-",
				temp.get("time")
				)
	html += "</table> </body>"
	with open(namespace.output + ".html", "w", encoding="UTF-16") as out:
		out.write(html)


def get_devices(f, section):
	item_re = re.compile(r"SYSTEM\\ControlSet\d*\\Enum\\" + section + r"\\([^\\^\s]+)\\([^\\^\s]+)\"$")
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
		# if other_section_re.search(line) != None:
		# 	break
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


if __name__ == "__main__":
	parser = create_parser()
	namespace = parser.parse_args()
	main(namespace)