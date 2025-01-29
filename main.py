from openpyxl import load_workbook
import pandas as pd


book = load_workbook('work.xlsx')
sheet = book.active
print("-> There is only 6 column!!")

try:
    print("-> Do you want to create an excel sheet or display excel sheet??")
    print("-> C for create and D for display")
    user_choice = input("-> ")

    if user_choice.lower() == "c":
        user_info = int(input("Enter how many column you need: "))
        items = int(input("enter number of Rows: "))

        if user_info >= 7:
            print("Out of range currently it will only takes 6 columns")

        else:

            for j, i in enumerate(range(user_info)):

                if i+1 == 1:
                    col1 = input(f"Enter column Name {j+1}: ")
                    sheet['A1'].value = col1
                    book.save('work.xlsx')
                    print(col1)

                    for p in range(items):
                        p += 1
                        sheet[f'A{p+1}'].value = input(f"{p}. ")
                        book.save('work.xlsx')

                elif i+1 == 2:
                    col1 = input(f"Enter column Name {j+1}: ")
                    sheet['B1'].value = col1
                    book.save('work.xlsx')
                    print(col1)

                    for p in range(items):
                        p += 1
                        sheet[f'B{p + 1}'].value = input(f"{p}. ")
                        book.save('work.xlsx')

                elif i+1 == 3:
                    col1 = input(f"Enter column Name {j+1}: ")
                    sheet['C1'].value = col1
                    book.save('work.xlsx')
                    print(col1)

                    for p in range(items):
                        p += 1
                        sheet[f'C{p + 1}'].value = input(f"{p}. ")
                        book.save('work.xlsx')

                elif i+1 == 4:
                    col1 = input(f"Enter column Name {j+1}: ")
                    sheet['D1'].value = col1
                    book.save('work.xlsx')
                    print(col1)

                    for p in range(items):
                        p += 1
                        sheet[f'D{p + 1}'].value = input(f"{p}. ")
                        book.save('work.xlsx')

                elif i+1 == 5:
                    col1 = input(f"Enter column Name {j+1}: ")
                    sheet['E1'].value = col1
                    book.save('work.xlsx')
                    print(col1)

                    for p in range(items):
                        p += 1
                        sheet[f'E{p + 1}'].value = input(f"{p}. ")
                        book.save('work.xlsx')

                elif i+1 == 6:
                    col1 = input(f"Enter column Name {j+1}: ")
                    sheet['F1'].value = col1
                    book.save('work.xlsx')
                    print(col1)

                    for p in range(items):
                        p += 1
                        sheet[f'F{p + 1}'].value = input(f"{p}. ")
                        book.save('work.xlsx')

                else:
                    print("invalid input!!")

    elif user_choice.lower() == "d":
        df = pd.read_excel('work.xlsx')
        print(df)

    else:
        print("!!Invalid Option!!")

except ValueError:
    print("ValueError Invalid information!!")