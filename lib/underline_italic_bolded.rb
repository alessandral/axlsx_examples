
require 'axlsx'
require 'pry'

if __FILE__ == $0
  puts "Spreadsheet with format bold, underline, italic"

  p = Axlsx::Package.new
  p.workbook do |w|

    w.styles do |st|

      format_underline  = st.add_style :u => true, :b => true, :i => true

      w.add_worksheet(:name => 'Sheet1') do |sheet|
        sheet.add_row ['Name', 'Address', 'Date', 'Amount'], :style => format_underline
        sheet.add_row ['Donald Duck', 'The tree top Street', Date.today.to_s, '100.00']
      end
    end

  end
  p.serialize 'sample_books/workbook.xlsx'

end
