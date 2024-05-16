from openpyxl.styles import Font, PatternFill
START_LETTER = 'A'
class Groups:
    def __init__(self, num_groups, participants):
        self.num_groups = num_groups
        self.num_participants = len(participants)
        # is group num valid
        self.group_size = self.num_participants // self.num_groups
        self.groups = []
        self.distribute_into_groups(participants)
        self.group_dictionary = {}
        self.create_dictionary()

    def distribute_into_groups(self, participants):
        self.groups = [participants[i * self.group_size:(i + 1) * self.group_size] for i in range(self.num_groups - 1)]
        self.groups.append(participants[(self.num_groups - 1) * self.group_size:])
    def is_group_num_valid(self):
        return self.num_participants % self.num_groups == 0

    def create_dictionary(self):
        self.group_dictionary = {
            'groups': [],
        }

        for ind, participants in enumerate(self.groups):
            group = {
                'num': chr(ord(START_LETTER) + ind),
                'participants': participants,
                'points': [0] * self.group_size
            }
            self.group_dictionary['groups'].append(group)
    def get_groups(self):
        return self.groups

    def print_groups(self, work_sheet):
        for i, group in enumerate(self.group_dictionary['groups']):
            row_num = i*(self.group_size+2)+1
            group_cell = work_sheet.cell(row=row_num, column=1)
            group_cell.value = 'Group ' + group['num']
            group_cell.font = Font(size=14, bold=True)
            title_cell1 = work_sheet.cell(row=row_num, column=11)
            title_cell1.value = 'Participants'
            title_cell1.font = Font(bold=True)
            title_cell2 = work_sheet.cell(row=row_num, column=12)
            title_cell2.value = 'Points'
            title_cell2.font = Font(bold=True)
            bye_cell_row_num = row_num + 1
            bye_cell_column_num = 2

            for j, entry in enumerate(group['participants']):
                work_sheet.cell(row=row_num+j+1, column=1).value = entry
                work_sheet.cell(row=row_num, column=j+2).value = entry
                work_sheet.cell(row=row_num+j+1, column=11).value = entry
                bye_cell = work_sheet.cell(row=bye_cell_row_num, column=bye_cell_column_num)
                bye_cell.fill = PatternFill(start_color="44546A", end_color="44546A", fill_type="solid")
                bye_cell_row_num = bye_cell_row_num + 1
                bye_cell_column_num = bye_cell_column_num + 1

        return work_sheet

