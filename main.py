from participants import Participants
from groups import Groups
from matches import Matches
from input_file import InputFile

# TODO: Write readme
# TODO: figure out the general user interactivity
file_path = 'entries.xlsx'
fl = InputFile(file_path)

if input("Is the file ready?") in ["Yes", "yes"]:
    if fl.is_sheet_name_exist("Participants"):
        prt = Participants(fl.entries)
        fxt = Matches()
        if prt.is_entry_num_valid():
            num_groups = int(input("How many groups?"))
            grp = Groups(num_groups, prt.list_participants)
            if grp.is_group_num_valid():
                if input("Do you want to randomize groups?") in ["Yes", "yes"]:
                    if fl.is_sheet_name_exist('Group Stage'):
                        ws1 = fl.wb['Group Stage']
                        fl.wb.remove(ws1)
                    ws1 = fl.wb.create_sheet(title='Group Stage')
                    ws1 = grp.print_groups(ws1)
                    if fl.is_sheet_name_exist('Matches'):
                        ws2 = fl.wb['Matches']
                        fl.wb.remove(ws2)
                    ws2 = fl.wb.create_sheet(title='Matches')
                    matches = fxt.create_matches(grp.groups)
                    ws2 = fxt.print_matches(ws2)
                else:
                    print("here")
                    if fl.is_sheet_name_exist('Group Stage'):
                        ws1 = fl.wb['Group Stage']
                        grp.read_groups(ws1)
                        grp.read_winners(ws1)
                        grp.print_points(ws1)
                        if fl.is_sheet_name_exist('Matches'):
                            ws2 = fl.wb['Matches']
                            fl.wb.remove(ws2)
                        ws2 = fl.wb.create_sheet(title='Matches')
                        matches = fxt.create_matches(grp.groups)
                        ws2 = fxt.print_matches(ws2)
                    else:
                        print("Please create group stage sheet.")
            else:
                print("Invalid number of groups, try again.")
        else:
            print("Invalid number of participants, try again.")
    else:
        print("Please create participants sheet.")
else:
    print("Try again")

fl.wb.save(file_path)

# TODO: adjust the alignment of the texts on cells
# TODO: save work sheet as internal variable for classes









