
require 'axlsx'

if __FILE__ == $0
  puts "Spreadsheet with a filter"

  p = Axlsx::Package.new
  p.workbook do |w|

    w.add_worksheet(:name => 'Sheet1') do |sheet|
      sheet.add_row ['Name', 'Address', 'Date', 'Amount']
      sheet.add_row ['Donald Duck', 'The tree top Street', Date.today.to_s, '100.00']
      sheet.add_row ['Mickey Mouse', 'Moon Av', Date.today.to_s, '80.00']
      sheet.auto_filter  = "A1:D3"
    end

  end
  p.serialize 'sample_books/workbook.xlsx'
end
