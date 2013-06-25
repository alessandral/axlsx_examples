require 'axlsx'
require 'pry'

if __FILE__ == $0
  sample_name = "Create a workbook with one spreadsheet"
  puts "#{sample_name}"

  p = Axlsx::Package.new
  p.workbook do |w|
    w.add_worksheet(:name => 'Sheet1') do |s|
      s.add_row ['Name', 'Address', 'Date', 'Amount']
      s.add_row ['Donald Duck', 'The tree top Street', Date.today.to_s, '100.00']
    end
  end
  p.serialize 'sample_books/workbook.xlsx'

end
