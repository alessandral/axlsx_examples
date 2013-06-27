
require 'axlsx'
require 'pry'

if __FILE__ == $0
  puts "Formula on a cell"

  p = Axlsx::Package.new
  p.workbook do |w|

    title = w.styles.add_style(:sz=>14,  :border=> {:style => :thin, :color => "FFFF0000"}, :alignment => {:horizontal => :center})
    style_money = w.styles.add_style(b: :true, num_fmt: 5)

    w.add_worksheet(:name => 'Sheet1') do |sheet|
      sheet.add_row ['Name', 'Address', 'Date', 'Amount'], :style => title
      sheet.add_row ['Donald Duck', 'The tree top Street', Date.today.to_s, '100.00']
      sheet.add_row ['Mickey', 'Sky blue Av', Date.today.to_s, '80.00']
      sheet.add_row [nil, nil,nil, "=SUM(D2:D3)"], :style => [nil, nil,nil, style_money]
    end

  end
  p.serialize 'sample_books/workbook.xlsx'
end

