require 'axlsx'
require 'pry'

if __FILE__ == $0
  puts "Pattern fill"

  p = Axlsx::Package.new
  p.workbook do |w|

    title = w.styles.add_style(:bg_color => "C0C0C0", :fg_color=>"#FF000000", :sz=>14,  :border=> {:style => :thin, :color => "FFFF0000"}, :alignment => {:horizontal => :center})

    pattern_style_light = w.styles.add_style({bg_color: '9999FF', b: true, type: :xf, border: {color: '9f9e9e', style: :thin}})
    pattern_style_dark  = w.styles.add_style({bg_color: 'CCCCFF', b: false, border: {color: '9f9e9e', style: :thin}})

    w.add_worksheet(:name => 'Sheet1') do |sheet|
      sheet.add_row ['Name', 'Address', 'Date', 'Amount'], :style => title
      sheet.add_row ['Donald Duck', 'The tree top Street', Date.today.to_s, '100.00'], :style => pattern_style_light
      sheet.add_row ['Mickey', 'Sky blue Av', Date.today.to_s, '80.00'], :style => pattern_style_dark
    end

  end
  p.serialize 'sample_books/workbook.xlsx'
end

