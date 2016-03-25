# Сохранение файла выгрузки данных из 1С в Excel

Экспортирует файл в несколько Excel-файлов по названию сущностей:

* Справочники
* Регистры сведений 

# Использование
usage: parseXML1CtoExcel.exe [-h] [--cache CACHE_FILE] input_file output_dir

positional arguments:
  input_file          выгруженный XML файл
  output_dir          выходная папка

optional arguments:
  -h, --help          показать помощь
  --cache CACHE_FILE  файл для создания кэша


## TODO
Сделать выгрузку документов
