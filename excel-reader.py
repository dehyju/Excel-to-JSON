import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import json
import os

current_directory = os.getcwd()

class ExcelConverterApp:
    def __init__(self,root):
        # 
        self.root = root
        self.selectedFile = ""
        # Creates window to select file to convert to JSON
        self.root.title("Excel to JSON converter")
        self.root.geometry("650x400")

        # Components
        self.lblTitle = tk.Label(self.root, text="Convert Excel to JSON", font=("Arial",25))
        self.lblSelectFile = tk.Label(self.root, text="Select File To Convert", font=("Arial",16))
        self.lblSubtitle = tk.Label(self.root, text="Press Button to Start")

        def convertFile():
            if self.selectedFile == "":
                self.lblConvertMessage.config(text="No file selected",fg="red")
            else:
                fileMade = self.makeToJson(self.selectedFile)
                if fileMade[1]:
                    self.lblConvertMessage.config(text="File created: "+fileMade[0],fg="green")
                    self.selectedFile = ""
                else:
                    self.lblConvertMessage.config(text="File already exists",fg="orange")
        self.select_button = tk.Button(self.root, text="Select Excel File", command=self.selectFile)
        self.btnConvertFile = tk.Button(self.root, text="Convert File", font=("Arial", 10), command=convertFile)

        self.lblConvertMessage = tk.Label(self.root, text="", font=("Arial",10))

        # Packing to window
        # Titles and subtitles
        self.lblTitle.pack(pady=10)
        self.lblSelectFile.pack(pady=50)
        self.lblSubtitle.pack(pady=10)

        # Place the dropdown in the window
        self.select_button.pack(pady=10)

        self.btnConvertFile.pack(pady=10)
        self.lblConvertMessage.pack()

    def selectFile(self):
        # Open file explorer and select an Excel file
        file_path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=(("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*"))
        )

        if file_path:  # If a file is selected
            self.lblConvertMessage.config(text=f"Processing {os.path.basename(file_path)}...",fg="black")
            self.selectedFile = file_path
            # self.process_excel(file_path)
        else:
            self.lblConvertMessage.config(text="No file selected. Please select an Excel file.",fg="red")
            self.selectedFile = ""

    def go_to_file(self):
        # Open file dialog and get the file path
        folder_path = filedialog.askdirectory(title="Select a Folder to Save JSON")
        
        if folder_path:  # If a file was selected
            print(f"Selected file: {folder_path}")
            # You can also do something with the file path here, like open it or read its contents
            return folder_path
        else:
            print("No file selected")  
            return ""
    # Converts file to JSON
    def makeToJson(self,excel_file): 

        # New JSON name
        json_basename, extension = os.path.splitext(excel_file)
        print(excel_file)
        json_filename = os.path.basename(json_basename).lower().replace(" ","_") + ".json"
        print(json_filename)

        savepath = self.go_to_file()
        savepath += "/" + json_filename

        if os.path.exists(savepath):
            response = messagebox.askquestion("Overwrite", "File already exists. Do you want to overwrite it?")
            if response == "yes":
                return self.convertFile(excel_file,savepath,json_filename)
            else:
                return [None,False]
        else:
            return self.convertFile(excel_file,savepath,json_filename)
            
            # print("Conversion complete! Check the", json_filename)
    def convertFile(self,excel_file,savepath,json_filename):
        # Read the Excel file into a DataFrame
        df = pd.read_excel(excel_file)

        # For NaN to Null conversion
        df = df.astype('object')

        # Replace NaN values with None
        df.replace('', None, inplace=True)
        df = df.apply(lambda col: col.where(pd.notnull(col), None))
        
        # Convert the DataFrame to a list of dictionaries (records)
        records = df.to_dict(orient='records')

        # Wrap the records in a list (array)
        json_data = json.dumps(records, indent=4)

        # Save the JSON data to a file
        with open(savepath, 'w') as json_file:
            json_file.write(f"{json_data}")

        return [json_filename,True]

def main():
    root = tk.Tk()

    # Initialize the file explorer app
    app = ExcelConverterApp(root)

    root.mainloop()

if __name__ == "__main__":
    main()
