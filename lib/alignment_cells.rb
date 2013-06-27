require 'axlsx'
require 'pry'

if __FILE__ == $0
  puts "Pattern fill"

  p = Axlsx::Package.new
  p.workbook do |w|

    title = w.styles.add_style(:bg_color => "333399", :fg_color=>"FFFFFF", :sz=>14,  :border=> {:style => :thin, :color => "000000"}, :alignment => {:horizontal => :center})

    cell_alignment = w.styles.add_style(:alignment => {:horizontal => :center, :vertical => :center, :wrap_text => :true})

    default_styling = {style: cell_alignment, widths: [nil,10], height: 60}

    w.add_worksheet(:name => 'Sheet1') do |sheet|
      sheet.add_row ['Name', 'Address', 'Date', 'Amount'], :style => title
      sheet.add_row ['Donald Duck', 'The tree top Street the street is so big in text that will need to wrap around the cell to fit all its content.', Date.today.to_s, '100.00'],
        default_styling.dup
      sheet.add_row ['Mickey', 'Sky blue Av', Date.today.to_s, '80.00'], default_styling.dup
    end

  end
  p.serialize 'sample_books/workbook.xlsx'
end

