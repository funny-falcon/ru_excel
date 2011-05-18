require 'ru_excel'

w = Excel::Workbook.new
ws = w.add_sheet('Multiline')

style = Excel::XFStyle.new
style.alignment.wrap = 1

ws.write(1, 1, "line 1 \n line A", style)

ws.col(1).width = 8000

File.open('multiline1.xls','w'){|f| w.save(f)}
