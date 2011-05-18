# -*- coding: utf-8 -*-
require 'ru_excel'

Excel.encoding = 'utf8'
wb = Excel::Workbook.new
ws = wb.add_sheet('info')
headers = ['Договор', 'Клиент', 'Дата', 'Оплачено',
                'Тип оплаты', 'Услуга']
headers.each_with_index{|name, i|
  ws.write(1, i, name)
}
File.open('utf8.xls','w'){|f| wb.save(f)}