require 'axlsx'
require 'pry'

if __FILE__ == $0
  puts "Protected Sheet"

  p = Axlsx::Package.new
  p.workbook do |w|
    w.add_worksheet(:name => 'Sheet1') do |sheet|

      sheet.sheet_protection.password = 'bananas'
      sheet.add_row ['Name', 'Address', 'Date', 'Amount']
      sheet.add_row ['Donald Duck', 'The tree top Street', Date.today.to_s, '100.00']
    end

  end
  p.serialize 'sample_books/workbook.xlsx'
end
