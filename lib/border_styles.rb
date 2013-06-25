
require 'axlsx'
require 'pry'

if __FILE__ == $0
  puts "Spreadsheet with border on a single row"

  p = Axlsx::Package.new
  p.workbook do |w|

    special_border = w.styles.add_style :border => {:style => :thin, :color =>"FAAC58"},  :border_bottom => {:style  => :thick }

    w.add_worksheet(:name => 'Sheet1') do |sheet|
      sheet.add_row ['Name', 'Address', 'Date', 'Amount'], :style => special_border
      sheet.add_row ['Donald Duck', 'The tree top Street', Date.today.to_s, '100.00']
    end

  end
  p.serialize 'sample_books/workbook.xlsx'
end
