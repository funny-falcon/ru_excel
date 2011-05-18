$: << '.'
require 'ru_excel'

style = Excel::XFStyle.new()

wb = Excel::Workbook.new()
ws0 = wb.add_sheet('0')

colcount = 200 + 1
rowcount = 800 + 1

t0 = Time.now
puts "\nstart: %s" % t0.to_s

puts "Filling..."
for col in 0...colcount
    puts "[%d]" % col if col % 10 == 0
    for row in 0...rowcount
        ws0.write(row, col, "BIG(%d, %d)" % [row, col])
        #ws0.write(row, col, "BIG")
    end
end

t1 = Time.now.to_f - t0.to_f
puts "\nsince starting elapsed %.2f s" % (t1)

puts "Storing..."
File.open('big-16Mb1.xls','wb')do|f|
    wb.save(f)
end

t2 = Time.now.to_f - t0.to_f
puts "since starting elapsed %.2f s" % (t2)


