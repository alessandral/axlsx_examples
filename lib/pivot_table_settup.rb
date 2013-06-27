require 'axlsx'
require 'pry'

if __FILE__ == $0
  puts "Pivot Table settup"

  def name
    ['Mickey', 'Donald', 'Clumsy', 'Patsy', 'Rusty', 'Toby'].sample
  end

  def product
    ['Blue glass', 'Yellow spoon', 'Red pencil', 'Brown dish', 'White napkin'].sample
  end

  def date
    ['2013-03-25', '2013-04-01'].sample
  end

  def amount
    ['80.00', '20.00', '10.00', '50.00'].sample
  end

  def place
    ['XYZ', 'PLK'].sample
  end

  p = Axlsx::Package.new
  p.workbook do |w|

    title = w.styles.add_style(:bg_color => "333399", :fg_color=>"FFFFFF", :sz=>14,  :border=> {:style => :thin, :color => "000000"}, :alignment => {:horizontal => :center})

    w.add_worksheet(:name => 'Sheet1') do |sheet|
      sheet.add_row ['Name', 'Product', 'Date', 'Amount', 'Place'], :style => title
      20.times{sheet.add_row [name, product, date, amount, place]}
      sheet.add_pivot_table 'G4:L17', 'A1:E21' do |pivot_table|
        #pivot_table.rows = ['Name', 'Product']
        #pivot_table.columns = ['Date']
        #pivot_table.data = ['Amount']
        #pivot_table.pages = ['Place']
      end

    end

  end
  p.serialize 'sample_books/workbook.xlsx'
end

