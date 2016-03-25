# Сохранение файла выгрузки данных из 1С в Excel

Парсит выгруженный обработчиком "Выгрузка и загрузка в XML" xml-файл и экспортирует в несколько Excel-файлов по названию сущностей:

* Справочники
* Регистры сведений 

## Использование

    parseXML1CtoExcel.exe [-h] [--cache CACHE_FILE] input_file output_dir
    или
    python parseXML1CtoExcel.py [-h] [--cache CACHE_FILE] input_file output_dir
    
    обязательные аргументы:
      input_file          выгруженный обработчиком "Выгрузка и загрузка в XML" xml-файл
      output_dir          выходная папка
    
    необязательные аргументы:
      -h, --help          показать помощь
      --cache CACHE_FILE  файл для создания кэша

## TODO
Сделать выгрузку документов
