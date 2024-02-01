require 'simple_xlsx_reader'

CC_SHEET_NAME = 'CC'
SALES_SHEET_NAME = 'Продажи'

HEADERS = {
  :NAME => 'Наименование',
  :PRICE => 'Цена',
  :DISCOUNT => 'Скидка',
  :COST => 'Себестоимость',
  :MANAGER => 'Имя продавца', 
  :PROFIT => 'Прибыль'
}

doc = SimpleXlsxReader.open('./docs/sales.xlsm')

def cc_sheet_to_hash(cc_sheet)
  items = []

  cc_sheet.rows.each(headers: true) do |row|
    next if row[HEADERS[:NAME]].nil? || row[HEADERS[:COST]].nil?
    items.push row
  end

  {
    :name => cc_sheet.name,
    :items => items
  }
end

def month_sheet_to_hash(month_sheet, cc)
  items_cc = []
  revenue_cc_total = 0
  cost_cc_total = 0
  
  items_no_cc = []
  revenue_no_cc_total = 0

  discount_total = 0

  month_sheet.rows.each(headers: true) do |row|
    next if row[HEADERS[:NAME]].nil? || row[HEADERS[:PRICE]].nil? 

    price = row[HEADERS[:PRICE]]
    discount = row[HEADERS[:DISCOUNT]] || 0

    discount_total += discount
    discounted = price - (price * discount / 100) 

    cc_item = cc[:items].detect { |item| item[HEADERS[:NAME]].eql? row[HEADERS[:NAME]] }
    
    if cc_item.nil? then
      revenue_no_cc_total += discounted

      item = row
      item[HEADERS[:COST]] = nil 
      items_no_cc.push row
    else
      cost_cc = cc_item[HEADERS[:COST]]
      cost_cc_total += cost_cc
      revenue_cc_total += discounted
      
      item = row
      item[HEADERS[:COST]] = cost_cc
      items_cc.push item
    end
  end

  {
    :name => month_sheet.name,
    :cc => {
      :items => items_cc,
      :count => items_cc.count,
      :revenue => revenue_cc_total,
      :profit => revenue_cc_total - cost_cc_total,
      :discount_avg => discount_total / items_cc.count
    },
    :no_cc => {
      :items => items_no_cc,
      :count => items_no_cc.count,
      :revenue => revenue_no_cc_total,
      :profit => nil,
      :discount_avg => discount_total / items_no_cc.count
    }
  }
end

def manager_cc_items(cc_items)
  manager_cc = []

  cc_items.each do |item|
    next if item[HEADERS[:MANAGER]].nil?
    manager_cc.push item
  end
  
  uniq_manager_cc = manager_cc.uniq { |item| item[HEADERS[:MANAGER]] } 

  uniq_manager_cc_profit = uniq_manager_cc.collect do |uniq_item|
    items_by_manager = manager_cc.select do |item|
      item[HEADERS[:MANAGER]].equal? uniq_item[HEADERS[:MANAGER]]
    end

    uniq_item[HEADERS[:PROFIT]] = items_by_manager.inject(0) do |sum, item|
      price = item[HEADERS[:PRICE]]
      discount = item[HEADERS[:DISCOUNT]] || 0
      discounted = price - (price * discount / 100)
      sum + (discounted - item[HEADERS[:COST]])
    end

    {
      HEADERS[:MANAGER] => uniq_item[HEADERS[:MANAGER]],
      HEADERS[:PROFIT] => uniq_item[HEADERS[:PROFIT]],
    }
  end
end

def cc_items_with_profit(cc_items)
  uniq_items = cc_items.uniq { |item| item[HEADERS[:NAME]] }
  
  uniq_items_profit = uniq_items.collect do |uniq_item|
    items_by_name = cc_items.select do |cc_item|
      cc_item[HEADERS[:NAME]].equal? uniq_item[HEADERS[:NAME]]
    end

    uniq_item[HEADERS[:PROFIT]] = items_by_name.inject(0) do |sum, item|
      price = item[HEADERS[:PRICE]]
      discount = item[HEADERS[:DISCOUNT]] || 0
      discounted = price - (price * discount / 100)
      sum + (discounted - item[HEADERS[:COST]])
    end

    {
      HEADERS[:NAME] => uniq_item[HEADERS[:NAME]],
      HEADERS[:PROFIT] => uniq_item[HEADERS[:PROFIT]]
    }
  end 
end

def cc_items_best_profit(cc_items)
  cc_items.sort_by { |item| item[HEADERS[:PROFIT]] }.reverse[0..99]
end

cc = nil
months = []

doc.sheets.each do |sheet|
  cc = cc_sheet_to_hash sheet if sheet.name.eql? CC_SHEET_NAME
end

doc.sheets.each do |sheet|
  months.push month_sheet_to_hash(sheet, cc) unless [CC_SHEET_NAME, SALES_SHEET_NAME].any? { |name| name.eql? sheet.name } 
end

months_total = months.collect do |month|
  {
    :name => month[:name],
    :count_cc => month[:cc][:count],
    :count_no_cc => month[:no_cc][:count],
    :revenue_cc => month[:cc][:revenue],
    :revenue_no_cc => month[:no_cc][:revenue],
    :profit_cc => month[:cc][:profit],
    :profit_no_cc => month[:no_cc][:profit],
    :discount_cc_avg => month[:cc][:discount_avg],
    :discount_no_cc_avg => month[:no_cc][:discount_avg],
    :cc_items_best_profit => cc_items_best_profit(cc_items_with_profit(month[:cc][:items])),
    :cc_items_best_manager => cc_items_best_profit(manager_cc_items(month[:cc][:items]))
  }
end

overall_cc_items = []
months.each { |month| overall_cc_items.push month[:cc][:items] }

overall_cc_items_best_profit = cc_items_best_profit(
  cc_items_with_profit(
    overall_cc_items.flatten
  )
)

overall_cc_items_best_manager = cc_items_best_profit(
  manager_cc_items(
    overall_cc_items.flatten
  )
)

# 
# WRITE
#

require 'rubyXL'
require 'rubyXL/convenience_methods'

workbook = RubyXL::Workbook.new

sheet_1 = workbook.add_worksheet('Продажи с номенклатурой')
sheet_1_row_1 = ['Месяц', 'Кол-во', 'Выручка', 'Прибыль', 'Средний % скидки']
sheet_1_row_1.each_with_index do |item, index|
  sheet_1.add_cell(0, index, item)
end
months_total.each_with_index do |month, index|
  sheet_1.add_cell(index + 1, 0, month[:name])
  sheet_1.add_cell(index + 1, 1, month[:count_cc])
  sheet_1.add_cell(index + 1, 2, month[:revenue_cc])
  sheet_1.add_cell(index + 1, 3, month[:profit_cc])
  sheet_1.add_cell(index + 1, 4, month[:discount_cc_avg])
end

sheet_2 = workbook.add_worksheet('Продажи без номенклатуры')
sheet_2_row_1 = ['Месяц', 'Кол-во', 'Выручка', 'Средний % скидки']
sheet_2_row_1.each_with_index do |item, index|
  sheet_2.add_cell(0, index, item)
end
months_total.each_with_index do |month, index|
  sheet_2.add_cell(index + 1, 0, month[:name])
  sheet_2.add_cell(index + 1, 1, month[:count_no_cc])
  sheet_2.add_cell(index + 1, 2, month[:revenue_no_cc])
  sheet_2.add_cell(index + 1, 3, month[:discount_no_cc_avg])
end

sheet_3 = workbook.add_worksheet('100 лучших позиций')
sheet_3_row_1 = [HEADERS[:NAME], HEADERS[:PROFIT]]
sheet_3_row_1.each_with_index do |item, index|
  sheet_3.add_cell(0, index, item)
end
overall_cc_items_best_profit.each_with_index do |item, index|
  sheet_3.add_cell(index + 1, 0, item[HEADERS[:NAME]])
  sheet_3.add_cell(index + 1, 1, item[HEADERS[:PROFIT]])
end

sheet_4 = workbook.add_worksheet('10 лучших менеджеров')
sheet_4_row_1 = [HEADERS[:MANAGER], HEADERS[:PROFIT]]
sheet_4_row_1.each_with_index do |item, index|
  sheet_4.add_cell(0, index, item)
end
overall_cc_items_best_manager.each_with_index do |item, index|
  sheet_4.add_cell(index + 1, 0, item[HEADERS[:MANAGER]])
  sheet_4.add_cell(index + 1, 1, item[HEADERS[:PROFIT]])
end

workbook.write('./result.xlsx')
puts 'done'
