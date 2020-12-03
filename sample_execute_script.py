# You'll need to modify the data_directory to point to the location you are dropping your files and your achive_directory
# to where you want to files to go after processing.


from budget_extraction import extract_budget_data

data_directory = 'C:\DOE Project\\'
archive_directory = 'C:\DOE Project\Archive\\'
server_name = 'localhost\SQLEXPRESS'
database_name = 'Budget'

extract_budget_data(data_directory,archive_directory)
