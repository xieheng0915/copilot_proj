require 'daru'
require 'write_xlsx'

# Function to return all Excel files in a directory
def get_excel_files(path)
  excel_files = []
  Dir.foreach(path) do |file|
    if file.end_with?('.xlsx')
      excel_files << file
    end
  end
  excel_files
end

# Path to directory containing Excel files
path = '/Users/xieheng/gitspace/python3space/copilot_with_py/data'

# Get all Excel files in the directory
excel_files = get_excel_files(path)

# Print the Excel files
puts excel_files

# Function to return dataframe from Excel path
def get_df_from_excel(path, sheet_name)
  df = Daru::DataFrame.from_excel(path, sheet: sheet_name)
  df
end

# Create a new folder "data_new" to store the new Excel files
new_path = '/Users/xieheng/gitspace/python3space/copilot_with_py/data_new'
Dir.mkdir(new_path) unless File.exists?(new_path)

# Get each dataframe from each Excel file
excel_files.each do |file|
  df = get_df_from_excel(File.join(path, file), 'Sheet1')
  
  # Replace each cell to 0 if the value is not a number
  df = df.map_rows { |row| row.map { |value| value.is_a?(Numeric) ? value : 0 } }
  
  # Sum each row
  df['Sum'] = df.row_vectors.map(&:sum)
  
  # Save the new dataframe to a new Excel file by the same name add "_new" to the end
  writer = WriteXLSX.new(File.join(new_path, File.basename(file, '.*') + '_new.xlsx'))
  df.write_excel(writer, sheet_name: 'Sheet1', index: false)
  writer.close
  
  puts df
end
