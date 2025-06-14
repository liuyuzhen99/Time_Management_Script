
class BiDirectionalMap:
    def __init__(self):
        self.index_2_staff = {}
        self.staff_2_index = {}

    def add_mapping(self, index, staff_code):
        self.index_2_staff[index] = staff_code
        self.staff_2_index[staff_code] = index

    def get_staff_by_index(self, index):
        return self.index_2_staff.get(index)

    def get_index_by_staff(self, staff_code):
        return self.staff_2_index.get(staff_code)
