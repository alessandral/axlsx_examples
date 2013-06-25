require 'axlsx'

if __FILE__ == $0
  puts "Format cells independiently"

  p = Axlsx::Package.new
  p.workbook do |w|

    format_bold      = w.styles.add_style :b => true
    format_italic    = w.styles.add_style :i => true
    format_underline = w.styles.add_style :u => true

    w.add_worksheet(:name => 'Sheet1') do |sheet|
      sheet.add_row ['Name', 'Address', 'Date', 'Amount']
      sheet.add_row ['Donald Duck', 'The tree top Street', Date.today.to_s, '100.00'], style: [format_bold, format_italic, format_underline]
      sheet.add_row ['Mickey', 'Flower land', Date.today.to_s, '80.00'], style: [format_italic, nil, nil]
    end

  end
  p.serialize 'sample_books/workbook.xlsx'
end
