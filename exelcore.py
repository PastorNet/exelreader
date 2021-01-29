import openpyxl


class XLSCORE:

    def __init__(self, exel_path: str):
        self.__exel = openpyxl.load_workbook(exel_path)
        self.__sheet_names = self.__exel.get_sheet_names()
        self.__sheets = {}
        for x_ in self.__sheet_names:
            self.__sheets[x_] = self.__exel.get_sheet_by_name(x_)

    @property
    def get_workbook_object(self) -> openpyxl.Workbook:
        return self.__exel

    @property
    def get_sheets(self) -> dict:
        return self.__sheets

    @property
    def get_sheets_names(self):
        return self.__sheet_names

    def get_sheet_by_key(self, sheet_key_: str):
        return self.__sheets.get(sheet_key_)

    @staticmethod
    def get_merged_cell_ranges(sheet):
        return sorted(sheet.merged_cell_ranges)

    def found_category_coordinate(self, correct_category_name: str, unmerged_sheet, axis=0) -> str:

        if axis:
            enumerable_sheet = unmerged_sheet.rows
        else:
            enumerable_sheet = unmerged_sheet.columns
        for x in enumerable_sheet:
            for y in x:
                if str(y.value) == correct_category_name:
                    return str(y.coordinate)

    @staticmethod
    def unmerge_sheet(sheet):
        [sheet.unmerge_cells(str(items)) for items in XLSCORE.get_merged_cell_ranges(sheet)]
        return sheet


class Schedule(XLSCORE):

    def __init__(self, course: str, group_name: str, shchedule_path: str):
        super().__init__(shchedule_path)
        self.__course = course
        self.__group = group_name
        self.__sheet = self.unmerge_sheet(self.get_sheet_by_key(self.__course))
        self.__group_coordinate = self.found_category_coordinate(self.__group, self.__sheet, axis=0)

    def get_pairs(self, current_day_name: str, last_day_name: str) -> list:
        letter_key = ''
        number_key = ''
        current_day_coordinate = self.found_category_coordinate(current_day_name, self.__sheet, axis=1)
        last_day_coordinate = self.found_category_coordinate(last_day_name, self.__sheet, axis=1)
        for x in self.__group_coordinate:
            if not str.isdigit(x):
                letter_key += x

        for x in last_day_coordinate:
            if str.isdigit(x):
                number_key += x
        last_key = letter_key + number_key
        pairs = []
        for x in self.__sheet[current_day_coordinate:last_key]:
            for y in x:
                if y.value is not None:
                    pairs.append(y.value)

        return pairs[0:len(pairs)-2]


if __name__ == '__main__':
    schedule = Schedule('1 курс', '101', 'testworkbook.xlsx')
    print(schedule.get_pairs('Вівторок', 'Четвер'))
