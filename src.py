#IMPORTANT!  DOWNLOAD EXCEL FILE(Book.xlsx) 
import sys 
import os 
from xlrd import open_workbook
from xlutils.copy import copy
import pandas as pd

#Opens the file
file = "/Users/prydatkomaryan/Desktop/Book.xlsx"

#Making and opening the Data Frame
df = pd.read_excel("/Users/prydatkomaryan/Desktop/unt/Book.xlsx")
df_without_na = df.fillna("Available") #changing NaN to 'Available'
df_without_na.index += 1 #makes data frame to start with 1 instead of 0
print(df_without_na) #printing out data Frame

#Opens Excel file
xl_file = r"Book.xlsx"
rb = open_workbook(xl_file)
wb = copy(rb)
my_sheet = wb.get_sheet(0)

#Input parent details,using whie loop
while True:
     parent_details = input("Enter your LAST name\n")
     if parent_details.isalpha(): #if details contains only letters
          # print(f"\nWelcome Mr/Mrs {parent_details}!")
          print("\n Welcome Mr/Mrs ")
          break #breaking the llop
     elif len(parent_details) < 1:
          print("You inserted only ONE character,NUMBER or SPECIAL SYMBOL \n Try again") #starting again
     else:
          print("Invalid input.Try to put only letters") #starting again

#Main part of the code
def main():
     #Asking about preferable day
     preferable_day = input("\nWhat is your  preferable day?\n1)Monday\n2)Tuesday\n3)Wednesday\nPress 1,2 or 3\n") 
     
     #Transform letter(1,2,3) to day, needed for excel navigation between rows
     if preferable_day == '1':
          preferable_day_str = "Monday"

     elif preferable_day == '2':
          preferable_day_str = "Tuesday"

     elif preferable_day == '3':
          preferable_day_str = "Wednesday"

     else:
          print("\nInvalid input,try again")
          main() #return to main
     
     #Choosing options what program should do
     message1 = input("\nPress to choose\nA) - Book Appointment\nB) - Cancel Appointment\nC) - Change time of the appointment\nA, B, or C\n")

     if message1 == "A":
          print("\nYou chose the option 'Book'")

          #Checking if {parent details} appointment exist already
          if df_without_na[preferable_day_str].str.contains(parent_details).any():
               print("\nYou have an appointment already") 
               main() #starting again
 
          else:
               #Asking for booking an appointment
               booking_msg = input("\nThere is no appointment,yet\nWould you like to book new one?(Yes/No)\n")
               if booking_msg == "Yes":
                    #Booking the appointment
                    def booker():
                         #Showing timetable of the day
                         print("\nShowing timetable of the day\n", df_without_na)
                         slot = input("\nWhat slot is suitable for you?(1-9)\n")
                         #Checking if in the slot that parent chose is available or not
                         if pd.isnull(df.iloc[int(slot), int(preferable_day)]):
                              #Writing parent details in the slot and day the chose
                              my_sheet.write(int(slot), int(preferable_day), parent_details)
                              wb.save(xl_file) #saving file

                              #Opens Data Frame agan
                              dfm = pd.read_excel("Book.xlsx")
                              dfm_without_na = dfm.fillna("Available") #Changing NaN to Available
                              dfm_without_na.index += 1 #Making Data Frame to start with 1 instead of 0

                              #Function for asking the question about the time and possible second child
                              def times():
                                   print(dfm_without_na) #Printing new Data Frame 
                                   comfortable = input("Is this time comfortable for you?(Yes/No)\n")
                                   if comfortable == "Yes":
                                        
                                        def children():
                                             children = input("Do you have any other  children?(Yes/No)\n")

                                             if children == "Yes":
                                                  main() #returns to the main,so parents can book again

                                             elif children == "No":
                                                  os.system("open " + file) #opens the excel file
                                             else:
                                                  print("Invalid input,try again")
                                                  children()# return to the function times
                                        children() #closing the function

                                   elif comfortable == "No":
                                        main() #return to main
                                        
                                   else:
                                        print("Invalid input,try again")
                                        times() #return to times
                              times() #Closing the function          
                 
                         else:
                              print("\nThis slot is booked already,try to book an appointment on new slot")
                              main() #return to main
                    booker() # closing the function

               elif booking_msg == "No":
                    print("\nStarting again")
                    main() #return to main
               else:
                    print("\nWrong input,try again")
                    main() #return to main
          
     elif message1 == "B":
          #Cancelling appointment
          print("\nYou chose the option 'Cancel'")
          #Checking if parents have an appointment already or not
          if df_without_na[preferable_day_str].str.contains(parent_details).any():
               print("\nCancelling your appointment'")

               #Searching through the sheet for parent details and deleting it
               for sheet in rb.sheets():
                    for rowindx in range (sheet.nrows):
                         row = sheet.row(rowindx)
                         for colindx, cell in enumerate(row):
                              if cell.value == parent_details:
                                   my_sheet.write(int(rowindx), int(colindx), "") #deleting the parent details
                                   wb.save(xl_file) #saves the file
               os.system("open " + file) #opens the Excel file
          else:
               print("\nThere is no appointment,yet.Trying again\n")
               main() #return to main
          
     elif message1 == "C":
          #Changing appointment
          print("\nYou chose the option 'Change'")
          print(df_without_na) #printing Data Frame
          #Checking for an appointment and changing it to different day/time
          if df_without_na[preferable_day_str].str.contains(parent_details).any():
               new_data_slot = input("\nWhat is your NEW preferable day?\n1)Monday\n2)Tuesday\n3)Wednesday\nPress 1,2 or 3\n")

               #Searching through the sheet
               for sheet in rb.sheets():
                    for rowindx in range (sheet.nrows):
                         row = sheet.row(rowindx)
                         for colindx, cell in enumerate(row):
                              if cell.value == parent_details:
                                   my_sheet.write(int(rowindx), int(colindx), "") #deleting previous appointment

               slot_changer = input("\nWhat NEW slot is suitable for you?(1-9)\n")
               my_sheet.write(int(slot_changer), int(new_data_slot), parent_details) #adding new appointment
               wb.save(xl_file) #saves the file
               os.system("open " + file) #Opens the Excel file
          else:
               print("\nThere is no appointment,yet.Trying again\n") #Doesn't have an appointmnet
               main() #return to main

     else:
          #If pressend something differend instead of A,B or C
          x = input("Invalid input.Do you want to start again?(Yes/No)")
          if x == "Yes":
               main() #return to main
          else:
               sys.exit(0) #exiting program
main() #closing the function