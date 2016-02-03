#!/usr/bin/python3

import sys
import argparse
import os
from struct import *
from datetime import datetime, timedelta

def create_parser():
	parser = argparse.ArgumentParser(
		prog="Prefetch",
		description="Программа выполняет сбор криминалистических данных из несжатых prefetch файлов ОС семейства Windows (кроме Win10) и вывод полученных данных в виде html документа.",
		epilog="Created by Canis (canis.ferox@yandex.ru)"
		)
	parser.add_argument("-d", "--directory", required=True, help="Директория содержащая prefetch файлы")
	parser.add_argument("-o", "--output", required=True, help="Имя файла для сохранения отчета (без расширения)")
	return parser


def main(namespace):
	init_dir = os.getcwd()
	os.chdir(namespace.directory)
	print("Поиск *.pf файлов в каталоге \"{}\"".format(os.getcwd()))	
	files = os.listdir()
	data = parse_files(os.getcwd() ,files)
	os.chdir(init_dir)
	generate_report(data, namespace.output)
	print("Готово. Файл отчета сохранен с именем \"{}/{}.html\"".format(os.getcwd(), namespace.output))


def generate_report(data, output):
	html  = "<head><title> Результат анализа файлов Prefetch</title></head>\n"
	html += "<body>\n <table border=\"1\" cellpadding=\"2\" cellspacing=\"0\">\n"
	html += "<tr> <th>№</th><th>Имя файла</th><th>Хэш</th><th>Количество запусков</th><th>Название программы</th><th>Размер pf файла в Б.</th><th>Дата последнего запуска</th> </tr>\n"
	n = 1
	for item in data.keys():
		temp = data[item]
		html += "<tr> <td>{}</td><td>{}</td>".format(n, item)
		html += "<td>{}</td><td>{}</td><td>{}</td><td>{}</td><td>{}</td> </tr>\n".format(temp["hash"], temp["count"], temp["filename"], temp["size"], temp["date"])
		n += 1
	with open(output + ".html", "w", encoding="UTF-16") as out:
		out.write(html)


def parse_files(root, files):
	root = root + "/"
	result = {}
	for _file in files:
		try:
			os.chdir(_file)
			os.chdir("..")
			continue
		except Exception:
			pass
		with open(root + _file, "rb") as f:
			temp = parse_file(f)
			if len(temp.keys()) > 0:
				result[_file] = temp
	return result


def parse_file(file_desk):
	binary = file_desk.read()
	version, signature = unpack("I4s", binary[0:8])
	if signature != b"SCCA":
		return {}
	pf_filesize, pf_filename, pf_hash = unpack("I60sI", binary[0x0C:0x50])
	pf_date, pf_exec_counter = 0, 0
	if version == 17:
		# XP, 2003
		pf_date = unpack("q", binary[0x78:0x80])[0]
		pf_exec_counter = unpack("I", binary[0x90:0x94])[0]
	if version == 23:
		# Vista, 7
		pf_date = unpack("q", binary[0x80:0x88])[0]
		pf_exec_counter = unpack("I", binary[0x98:0x9C])[0]
	if version == 26:
		# 8.1
		pf_date = unpack("q", binary[0x80:0x88])[0]
		pf_exec_counter = unpack("I", binary[0xD0:0xD4])[0]
	if version == 30:
		# 10
		return {}
	return({
			"size" : pf_filesize,
			"filename" : pf_filename.decode("UTF-16"),
			"hash" : hex(pf_hash),
			"date" : get_date(pf_date),
			"count" : pf_exec_counter
			})


def get_date(win_binary_date):		# Преобразование бинарной даты Windows в обычную
	return datetime(1601,1,1) + timedelta(microseconds =  (win_binary_date / 10))


if __name__ == "__main__":
	parser = create_parser()
	namespace = parser.parse_args()
	main(namespace)