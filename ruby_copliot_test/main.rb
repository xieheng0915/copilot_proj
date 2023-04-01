
# write a function to return all excel files path in a directory  
def get_excel_files_path(dir_path)  
  excel_files = []  
  Dir.foreach(dir_path) do |file|  
  excel_files << File.join(dir_path, file) if file =~ /.xls$/  
  end  
  excel_files  
end

puts path = "/Users/xieheng/gitspace/ruby_space/ruby_copliot_test/data"

# Get all excel files in the directory
excel_files = get_excel_files_path(path)

# print excel files path
puts excel_files

# return dataframe from excel path
def get_dataframe_from_excel(excel_path)  
  df = Daru::DataFrame.from_excel(excel_path)  
  df  
end

# create a new folder "data_new" to store new excel files if folder doesn't exist 
new_path = "/Users/xieheng/gitspace/ruby_space/ruby_copliot_test/data_new"  
Dir.mkdir(new_path) unless File.exist?(new_path)

# get each dataframe from excel files
excel_files.each do |excel_file|  
  df = get_dataframe_from_excel(excel_file)  
  # Replace each cell to 0 if the value is not a number
  df.each_row do |row|  
    row.each do |cell|  
      cell.replace(0) unless cell.is_a?(Numeric)  
    end  
  end  
  # sum each row
  df = df.sum(axis: 1)

  # generate a new excel file, save the new dataframe to the new excel file by the same name add "_new" suffix
  new_excel_file = File.join(new_path, File.basename(excel_file, ".xls") + "_new.xls")
  
  df.to_excel(new_excel_file)


end