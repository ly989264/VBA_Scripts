# global variables

# revenue, modify if needed
arr = ["Total Job Active Services Revenue",
     "Total DES Revenue",
     "Total CTADD Revenue",
     "Total CTA Revenue",
     "Total TTW Revenue",
     "Total AASN Revenue",
     "Total SSH Revenue",
     "Total Supplementary Services Revenue",
     "Total Fee for Service Revenue",
     "Total Other Revenue"]
column = [11, 20]  # means that the revenue matching rows from 11 to 20, modify if needed

# expenses, modify if needed
arr2 = ["Total Job Active Services Expenses",
      "Total DES Expenses",
      "Total CTADD Expenses",
      "Total CTA Expenses",
      "Total TTW Expenses",
      "Total AASN Expenses",
      "Total SSH Expenses",
      "Total Supplementary Services Expense",
      "Total Fee for Service Expense",
      "Total Other Expenses"]
column2 = [24, 33]  # means that the revenue matching rows from 24 to 33, modify if needed

# expenses, modify if needed
arr3 = ["Total Staffing Expenses",
      "Total Travel & Accommodation  Expenses",
      "Total Office Accommodation Expenses",
      "Total General Meeting Expenses",
      "Total Data & Communication Expenses",
      "Total Marketing Expenses",
      "Total IT Expenses",
      "Total Legal & professional Expenses",
      "Total Financial & Insurance Expenses",
      "Total General Office Expenses",
      "Total Board Expenses",
      "Total Depreciation & Amortisation"]
column3 = [50, 61]  # means that the revenue matching rows from 50 to 61, modify if needed

# matching rules
# 1:'b' means that offset(0,1) should be copied to the corresponding column b
# 2:'c' means that offset(0,2) should be copied to the corresponding column c
# and so on
# modify if needed
a = {1:'b',2:'c',6:'g',7:'h',11:'l',12:'m'}

# paste special parameter, modify if needed
paste_special_params = "Paste:=xlPasteValues"

# number of indent spaces ahead, modify if needed
num_spaces = 4
spaces = " " * num_spaces  # automatically generated, do not modify


# generator function, do not modify
def generator(num):
    global a
    for i in sorted(a.keys()):
        print(f'{spaces}    current_value.Offset(0, {i}).Copy')
        # print(f'Worksheets(\"Consolidated\").select')
        print(f'{spaces}    Worksheets(\"Consolidated\").Range("{a[i]}{num}").PasteSpecial {paste_special_params}')


# generate procedure, do not modify
index = 0
print(f"{spaces}set current_value = Worksheets(\"Detailed\").Cells(i, 1)")
for i in range(column[0], column[1]+1):
    if index == 0:
        print(f"{spaces}If current_value.Value = \"{arr[index]}\" Then")
    else:
        print(f"{spaces}ElseIf current_value.Value = \"{arr[index]}\" Then")
    index += 1
    generator(i)
index = 0
for i in range(column2[0], column2[1]+1):
    print(f"{spaces}ElseIf current_value.Value = \"{arr2[index]}\" Then")
    index += 1
    generator(i)
index = 0
for i in range(column3[0], column3[1]+1):
    print(f"{spaces}ElseIf current_value.Value = \"{arr3[index]}\" Then")
    index += 1
    generator(i)
print(f"{spaces}End If")


# VBA code generator
