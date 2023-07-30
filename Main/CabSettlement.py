from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.dropdown import DropDown
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from datetime import datetime, timedelta
from ExcelData import ExcelData
import openpyxl


class startscreen(Screen):
    def __init__(self, **kwargs):
        super(startscreen, self).__init__(**kwargs)
        self.dropdown = DropDown()

        # Add numbers 1 to 10 to the dropdown
        for i in range(1, 101):
            btn = Button(text=str(i), size_hint_y=None, height=44)
            btn.bind(on_release=lambda btn: self.dropdown.select(btn.text))
            self.dropdown.add_widget(btn)

        self.main_layout = BoxLayout(orientation='vertical')
        self.name_input = TextInput(hint_text="Enter the starting date", multiline=False)
        self.btn_select_number = Button(text="No of Days Stay")
        self.btn_select_number.bind(on_release=self.dropdown.open)
        self.dropdown.bind(on_select=self.on_dropdown_select)
        self.btn_start_game = Button(text="Start Cab Fill", on_press=self.start_game)

        self.btn_end_game = EndGameButton(text="Close It!")

        self.main_layout.add_widget(self.name_input)
        self.main_layout.add_widget(self.btn_select_number)
        self.main_layout.add_widget(self.btn_start_game)
        self.main_layout.add_widget(self.btn_end_game)
        self.add_widget(self.main_layout)

        self.final_list = []

    def cabSettlementFromTo(self):
        start_date = '2023-07-26'
        #start_date = self.date_entry.get(1.0, tk.END)
        if len(start_date)<=0:
            start_date = '2023-07-26'
        #noOfDaysStay = 1
        noOfDaysStay = self.days_dropdown.get()
        self.date_plus_100_days(start_date,int(noOfDaysStay))


    def on_dropdown_select(self, instance, text):
        # Get the selected number from the dropdown
        self.selected_number = int(text)

    def start_game(self, instance):
        self.count = self.selected_number
        self.manager.current = "iscabbookedscreen"
        iscabbookedscreen = self.manager.get_screen("iscabbookedscreen")
        self.date_text = self.name_input.text


        if(self.date_text==""):
            self.date_text = "2023-07-30"

        self.manager.shared_date = datetime.strptime(self.date_text, '%Y-%m-%d')

        #self.cab_date = datetime.strptime(start_date, '%Y-%m-%d')
        #print("cab_date : "+self.cab_date)
        #iscabbookedscreen.set_cab_date(self.cab_date)
        iscabbookedscreen.set_count(self.count)


class EndGameButton(Button):
    def on_press(self):
        App.get_running_app().stop()


class iscabbookedscreen(Screen):
    def __init__(self, **kwargs):
        super(iscabbookedscreen, self).__init__(**kwargs)
        #current_date = self.manager.shared_date

        print("Reloaded iscabbookedscreen")

        self.date_text = "2023-07-07"

        self.date1 = datetime.strptime(self.date_text, '%Y-%m-%d')
        self.loop_times = 2

        self.morEven = "Morning" if self.loop_times%2==0 else "Evening"
        self.hotelOrOffice = "Taxi(Hotel to Office)" if self.loop_times%2==0 else "Taxi(Office to Hotel)"

        self.proceedWithDate = self.date1.strftime('%d/%b/%y')
        print("Shared date: " + self.proceedWithDate)

        self.main_layout = BoxLayout(orientation='vertical')
        self.label_name = Button(text=f"Is cab booked on {self.proceedWithDate}, {self.morEven} ?", size_hint_y=None, height=44)
        self.btn_yes = Button(text="Yes", on_press=self.on_yes_click)
        self.btn_no = Button(text="No", on_press=self.on_no_click)

        self.btn_main_menu = Button(text="Main Menu!", on_press=self.back_to_main_menu)

        #self.cab_date = startscreen.start_date


        self.main_layout.add_widget(self.label_name)
        self.main_layout.add_widget(self.btn_yes)
        self.main_layout.add_widget(self.btn_no)
        self.main_layout.add_widget(self.btn_main_menu)
        self.add_widget(self.main_layout)

        self.final_list = []  # List to store guessed numbers 

    def back_to_main_menu(self, instance):
        self.generate_excel(self.final_list)
        self.manager.current = "startscreen"

    def set_count(self, count):
        self.count = count

    def set_cab_date(self, date):
        self.cab_date = date

    def on_no_click(self, instance):
        print("No Guessed")
        #self.cab_date = self.cab_date + timedelta(days=1)
        #end_date = start_date + timedelta(days=no_of_days)


        self.loop_times -= 1

        if self.loop_times==0:
            self.date1 += timedelta(days=1)
            self.proceedWithDate = self.date1.strftime('%d/%b/%y')
            self.morEven = "Morning" if self.loop_times%2==0 else "Evening"
            self.hotelOrOffice = "Taxi(Hotel to Office)" if self.loop_times%2==0 else "Taxi(Office to Hotel)"
            self.label_name.text = f"Is cab booked on {self.proceedWithDate}, {self.morEven} ?"
            self.loop_times = 2
            self.count -= 1
        else:
            self.proceedWithDate = self.date1.strftime('%d/%b/%y')
            self.morEven = "Morning" if self.loop_times%2==0 else "Evening"
            self.hotelOrOffice = "Taxi(Hotel to Office)" if self.loop_times%2==0 else "Taxi(Office to Hotel)"
            self.label_name.text = f"Is cab booked on {self.proceedWithDate}, {self.morEven} ?"


        row = ExcelData(self.proceedWithDate, self.hotelOrOffice, "{:.2f}".format(0), "{:.2f}".format(1), "{:.2f}".format(0), "","NO")
        self.final_list.append(row)

        #self.manager.shared_date += timedelta(days=1)
        
        if self.count > 0:
            self.manager.current = "iscabbookedscreen"
        else:
            self.generate_excel(self.final_list)
            self.manager.current = "startscreen"
            #for lst in self.final_list:
            #    print(lst)
            #print(self.final_list)

        

    def on_yes_click(self, instance):
        startscreen = self.manager.get_screen("startscreen")
        #self.cab_date = self.cab_date + timedelta(days=1)

        self.loop_times -= 1

        if self.loop_times==0:
            self.date1 += timedelta(days=1)
            self.proceedWithDate = self.date1.strftime('%d/%b/%y')
            self.morEven = "Morning" if self.loop_times%2==0 else "Evening"
            self.hotelOrOffice = "Taxi(Hotel to Office)" if self.loop_times%2==0 else "Taxi(Office to Hotel)"
            self.label_name.text = f"Is cab booked on {self.proceedWithDate}, {self.morEven} ?"
            self.loop_times = 2
            self.count -= 1
        else:
            self.proceedWithDate = self.date1.strftime('%d/%b/%y')
            self.morEven = "Morning" if self.loop_times%2==0 else "Evening"
            self.hotelOrOffice = "Taxi(Hotel to Office)" if self.loop_times%2==0 else "Taxi(Office to Hotel)"
            self.label_name.text = f"Is cab booked on {self.proceedWithDate}, {self.morEven} ?"


        #self.manager.shared_date += timedelta(days=1)
        #print("Shared date: " + self.manager.shared_date)

        print("Yes, Guess")
        self.manager.current = "entercabamountscreen"
        if self.count > 0:
            entercabamountscreen.proceedWithDate = self.proceedWithDate
            entercabamountscreen.hotelOrOffice = self.hotelOrOffice
            self.manager.current = "entercabamountscreen"
        else:
            self.generate_excel(self.final_list)
            self.manager.current = "startscreen"
            #for lst in self.final_list:
            #    print(lst)
            #print(self.final_list)


    def generate_excel(self, excelDataLst):
        dataofAllRow = []

        for i in range(len(excelDataLst)):
            dataOfRow = []
            if excelDataLst[i].booked=="YES":
                dataOfRow.append(excelDataLst[i].date)
                dataOfRow.append(excelDataLst[i].hotelOffice)
                dataOfRow.append(excelDataLst[i].amount)
                dataOfRow.append(excelDataLst[i].exchRate)
                dataOfRow.append(excelDataLst[i].amountFC)
                dataOfRow.append(excelDataLst[i].remarks)
                #dataOfRow.append(excelDataLst[i].booked)
                
                dataofAllRow.append(dataOfRow)

        self.saveExcel(dataofAllRow)


    def saveExcel(self,dataofAllRow):
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # Add headers to the sheet
        #sheet.append(['Date','Paticulars(From-To)' , 'Amount', 'Exch. Rate', 'Amount(FC)','Remarks(with bill/without bill)', 'Booked'])
        sheet.append(['Date','Paticulars(From-To)' , 'Amount', 'Exch. Rate', 'Amount(FC)','Remarks(with bill/without bill)'])

        # Add sample data rows
        #data = [
         #   ['Alice', 30, 50000.0],
          #  ['Bob', 25, 60000.0],
           # ['Charlie', 35, 55000.0]
        #]

        for row in dataofAllRow:
            sheet.append(row)

        # Save the workbook to a file
        workbook.save('FinalExcel.xlsx')
        print("Excel file generated successfully!")






class entercabamountscreen(Screen):
    def __init__(self, **kwargs):
        super(entercabamountscreen, self).__init__(**kwargs)
        self.main_layout = BoxLayout(orientation='vertical')
        self.label_question = Button(text="Enter the amount:", size_hint_y=None, height=44)
        self.number_input = TextInput(hint_text="Enter amount", multiline=False)
        self.btn_ok = Button(text="Bolt", on_press=self.on_bolt_click)
        self.btn_cancel = Button(text="Uber", on_press=self.on_uber_click)

        self.main_layout.add_widget(self.label_question)
        self.main_layout.add_widget(self.number_input)
        self.main_layout.add_widget(self.btn_ok)
        self.main_layout.add_widget(self.btn_cancel)
        self.add_widget(self.main_layout)

    def on_bolt_click(self, instance):
        entered_amount = self.number_input.text
        print(f"Bolt : {entered_amount}")

        # Get the entered number from the TextInput before moving to iscabbookedscreen
        amount_decimal = "{:.2f}".format(float(entered_amount))
        #self.finalCabDetailsLst.append(int(entered_amount))
        
        row = ExcelData(self.proceedWithDate, self.hotelOrOffice, amount_decimal, "{:.2f}".format(1), amount_decimal, "BOLT","YES")
        
        #self.final_list.append(float(entered_amount))
        self.number_input.text = ""  # Reset the TextInput
        self.manager.current = "iscabbookedscreen"
        iscabbookedscreen = self.manager.get_screen("iscabbookedscreen")
        iscabbookedscreen.final_list.append(row)  # Add the guessed number to iscabbookedscreen's list


    def on_uber_click(self, instance):
        entered_amount = self.number_input.text
        print(f"Uber : {entered_amount}")

        amount_decimal = "{:.2f}".format(float(entered_amount))
        
        row = ExcelData(self.proceedWithDate, self.hotelOrOffice, amount_decimal, "{:.2f}".format(1), amount_decimal, "UBER","YES")
        
        #self.final_list.append(float(entered_amount))
        self.number_input.text = ""  # Reset the TextInput
        self.manager.current = "iscabbookedscreen"
        iscabbookedscreen = self.manager.get_screen("iscabbookedscreen")
        iscabbookedscreen.final_list.append(row)  # Add the guessed number to iscabbookedscreen's list


class CabSettlementApp(App):
    def build(self):
        sm = ScreenManager()

        sm.add_widget(startscreen(name="startscreen"))
        sm.add_widget(iscabbookedscreen(name="iscabbookedscreen"))
        sm.add_widget(entercabamountscreen(name="entercabamountscreen"))

        return sm


if __name__ == "__main__":
    CabSettlementApp().run()



'''

pyinstaller CabSettlement.py --onefile --windowed


pyinstaller main.py --onefile --windowed --add-data "path/to/kivy/data;./kivy/data

/Users/ardanish/Documents/A_R_Projects/venv/lib/python3.11/site-packages/kivy/data

pyinstaller main.py --onefile --windowed --add-data "/Users/ardanish/Documents/A_R_Projects/venv/lib/python3.11/site-packages/kivy/data;./kivy/data"

'''
