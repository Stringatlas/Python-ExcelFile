import os

try:
	import openpyxl
except:
	print("could not find pyxl module")



def main():
	current_path = os.getcwd()
	directory = "" # input("Directory: ")
	excelFile_Paths = []
	fileNames = []
	for file_in_directory in os.listdir(os.path.join(current_path, directory)):
		if file_in_directory.endswith(".xlsx"):
			excelFile_Paths.append(os.path.join(current_path, file_in_directory))
			fileNames.append(file_in_directory)

	with open("output.txt", "w") as outputFile:
		for fileIndex in range(0, len(excelFile_Paths)):
			excelFile = openpyxl.load_workbook(excelFile_Paths[fileIndex])
			for sheetname in excelFile.sheetnames:		
				outputFile.write(f"{fileNames[fileIndex]}: {sheetname}\n")

if __name__ == "__main__":
	main()