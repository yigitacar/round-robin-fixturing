class Participants:
    def __init__(self, participants):
        self.participants = participants
        self.shuffled_participants = self.shuffle_participants()
        self.list_participants = self.convert_into_list()
        self.num_participants = len(participants)
    def shuffle_participants(self):
        return self.participants.sample(frac=1).reset_index(drop=True)
    def convert_into_list(self):
        return self.shuffled_participants['Participants'].tolist()
    def is_entry_num_valid(self):
        return (self.num_participants & (self.num_participants - 1) == 0) and self.num_participants != 0