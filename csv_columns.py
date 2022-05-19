import os
import csv


current_path = os.getcwd()
directory = "" 
csvFile_Paths = []
fileNames = []

for file_in_directory in os.listdir(os.path.join(current_path, directory)):
	if file_in_directory.endswith(".csv"):
		csvFile_Paths.append(os.path.join(current_path, file_in_directory))
		fileNames.append(file_in_directory)




with open("columns.txt", "w") as columns_file:
	columns_file.write("File: Column\n")
	
	with open("paired_sensor_drop_validation.txt", "w") as validation_file:
		validation_file.write("File: Paired sensor/drop validation\n\n")


		for fileIndex in range(0, len(csvFile_Paths)):
			
			with open(csvFile_Paths[fileIndex], "r") as csvFile:
				no_pairs = []
				has_pair = False
				num_pairs = 0


				reader = csv.reader(csvFile)
				column_names = next(reader)
	
				for i in range(0, len(column_names)):
					column_names[i] = column_names[i].replace("\n", " ")
					column_name = column_names[i]
					columns_file.write(f"{fileNames[fileIndex]}: {column_name}\n")
					
					# validation file
					if has_pair:
						num_pairs += 1
						has_pair = False
					
					elif "censor" in column_name:
						model = column_name[0:-7]
						
						if model + " # of drops" == column_names[i+1]:
							has_pair = True
						else:
							no_pairs.append(i+1)
	
					elif "# of drops" in column_name:
						model = column_name[0:-11]
						
						if model + " censor" == column_names[i-1]:
							has_pair = True
						else:
							no_pairs.append(i-1)
	
				if not no_pairs:
					validation_file.write(f"{fileNames[fileIndex]}: True, Number of Pairs = {num_pairs}\n")
				else:
					validation_file.write(f"{fileNames[fileIndex]}: False, Number of Pairs = {num_pairs};\n")
					for no_pair in no_pairs:
						validation_file.write(f"column {no_pair+1}: {column_names[no_pair]}\n")
					
				validation_file.write("\n\n")





	


		


