import  openpyxl as xl
import pprint
import os, shutil
from openpyxl import Workbook

def load(name="data"):
    wb = xl.load_workbook("%s"%name)
    return wb
    
def create(wb, name):
    cwd = os.getcwd()
    os.makedirs(os.path.join(cwd, "%s"%(name)))
    os.chdir(cwd+"\\"+name)
    wb.save("%s.xlsx"%name)
def save(my_wb):
    ''' 
    Helper function to save the current workbook
    Parametre:
        - Receive ann WorkBook object
    '''
    my_wb.save(filename = "data.xlsx")

def openFile(name="", flag=""):
    ''' 
        Function to open the file. 
        Parametres:
            - A string called source. The name of the file (by default it is source).
    '''
    if flag == "r":
        name = "entries"
    elif flag == "c":
        name = "columns"
    file = open("%s.txt"%name, "w+")
    return file

def openExists(filename):
    cwd = os.getcwd()
    try:
        os.chdir(cwd +"\%s"%file)
    except FileNotFoundError:
        raise "File not found"
    return  open("columns.txt", "r+")
def show_info():
    print("""
    This is an utility script, it allows you to fill empyties 
    excel documents without open then.
    """)

def add_new_field(file, fields_quantity):
   
    for i in range(fields_quantity):
        new_field = input("Enter the Field(Like: Name, number, address etc): ")
        file.write(new_field+"\n")
    file.close()
def show_menu():
    print("="*52)
    print("="*4, " "*7, "Welcome to Txt to Excel v0", " "*7, "="*4)
    print("="*4, "By Antonio Pedro", " "*25, "="*4)
    print("="*52)
    print("\t"*5)
    print("="*23, "Menu", "="*23)
    print("Press: \n1. To create a new document\t\t2. List all documents")
    print("3. To get help\t\t\t\t4. Exit")
if __name__ == '__main__':
    show_menu()
    option = int()
    excel_filename = ""
    index = 0
    wb = Workbook()
    while option != 4:
        option = int(input(">>> "))
        if option == 1:
            excel_filename = input(f"Enter The name of the excel file(like: payments, my_plan etc): ")
            try:
                create(wb, excel_filename)
                fields = openFile(name="columns")
                entries = openFile(name="entries")
                saver = open("save.bat", "w+")
                entries.write(""""
                ==========================================================
                This is an utility script, it allows you to fill empyties 
                    excel documents without open then''.
                    Fill the entries as in the fields orders and make
                    sure to add '-' once entered new record. Your final document should looks like:

                    0
                    Antonio
                    100
                    Good
                    -
                    1
                    Isabel
                    100
                    Good
                    -
                    2
                    Alves
                    90
                    Good
                    -
                    3
                    Luis
                    54
                    Avarage
                    -
                    NOTE: PLEASE DELETE ALL THESE LINES BEFORE STARTING.
                """"")
                current_path = os.getcwd()
                previous = current_path[0:current_path.index(excel_filename)] + "config.pyw"
                shutil.copy(previous, current_path)
                saver.write(f"@echo off\ncd ..\npython config.pyw run -b")
            
            except BaseException:
               print("Sorry We could not create you file")

            print("""Note: Once created a new excel document, there are two other Text :
                1. Columns.txt - Where you will enter the desired fields of your excel document
                1. Entries.txt - Where you will enter the desired informations of your excel document""")
            fields_quantity = int(input("How many fields do you want to add? "))
            add_new_field(fields, fields_quantity)
            print("Everything is done now!\nGo to ", os.getcwd(), "folder\nOpen Entries.txt file and start add entries\nAfter filled, run the file save.bat")
            opt = input("Press 0 to exit or any key to continue: ")
            if opt== '0':
                break
            else:
                if index > 0:
                    os.chdir(os.getcwd()[0:os.getcwd().index(excel_filename)])
                print(os.getcwd())
                show_menu()
        elif option == 2:
            print("Bellow are all documents you have: ")
            for file in os.listdir():
                if os.path.isdir(file):
                    print(f"{index+1}. ", file)
                    index = index + 1
            opt = input("Press 0 to exit or any key to continue: ")
            if opt == '0':
                break
            else:
                if index > 0:
                    os.chdir(os.getcwd()[0:os.getcwd().index(excel_filename)])
                show_menu()
        elif option == 3:
            show_info()
            opt = input("Press 0 to exit or any key to continue: ")
            if opt == '0':
                break
            else:
               if index > 0:
                    os.chdir(os.getcwd()[0:os.getcwd().index(excel_filename)])
               show_menu()
        elif option == 4:
            break
        else:
            print("Invalid choice")