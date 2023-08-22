# Общее описание

Проект предназначен для автоматизации проверки drawio диаграм, использующих стандарт C4, а так же созданеи модели по диаграмме в двух форматах:
- xls
- Structurizr DSL

Реализованы следующие проверки:

* Заполненность описаний компонент
* Заполненность технологий компонент
* Заполненность описаний связей.
  С*вязи изображают вызовы между компонентами.  На связях между компонентами указываются данные которые передаются от компонента к компоненту.** **
  Используется следующая нотация:* *название действия (передаваемые данные): возвращаемые данные [технологии]
  Например* *”Зарегистрировать заказ (абонент, продукт): заказ [*  *gRPC* *]»*
* Заполненность технологий связей

поддерживается как сжатый так и обычный формат drawio

Скрипит пытается исправить следующие проблемы в фалйах:

* Когда стрелки не присоеденены к объекту а только касаются
* Когда стрелки не являются стрелаками в нотации C4

# Инсталяция

* Для работы потребуется python3
* Потребуется установка xlsxwriter "pip3 install xlsxwriter""


# Использование

drawio_parser.py -i `<inputfile>` -o `<outputfile> -d -s`

inputfile - имя файла в формате drawio

outputfile - имя файла в формате xlsx в который будут записаны объекты и связи

d - проверка синтаксиса входных и выходных данных

s - печать статистики
