require 'backports/2.0'
require 'daru'
require 'daru/io/importers/excelx'
require 'reportbuilder'
require 'write_xlsx'
require 'spreadsheet'
require 'roo'


# Read data from Excel file
xlsx = Daru::IO::Importers::Excelx.read('data.xlsx')
# data = xlsx[:sheet1]
data = xlsx.call(worksheet: 'Sheet1')

# Create a new Excel file and write data to it
workbook = Daru::DataFrame.new('output.xlsx')
worksheet = workbook.add_worksheet

# Write headers to worksheet
data.vectors.each_with_index do |header, index|
  worksheet.write(0, index, header)
end

# Write data to worksheet
data.each_row_with_index do |row, row_index|
  row.each_with_index do |value, col_index|
    worksheet.write(row_index+1, col_index, value)
  end
end

workbook.close

