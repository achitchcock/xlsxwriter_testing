import random
import string
import xlsxwriter  # type: ignore
from typing import List


class DemoWorkbook(object):
    def __init__(self, filename):
        # type: (str) -> None
        self.workbook = xlsxwriter.Workbook(filename + ".xlsx", None)
        self.worksheet = None
        self.write_times_table(65.0, 40.0)  # type: (float,float)
        self.roll_dice(10, 30)  # type: (int, int)
        self.itc_test()
        self.write_words()
        self.workbook.close()

    def write_times_table(self, x_size, y_size):
        # type: (float,float) -> None
        """
        Generates a times table from 0*0 to X*Y with a color gradient based on
        :param x_size: NUmber of columns to fill in the horizontal direction
        :param y_size: Number of columns to fill in the vertical direction
        :return: None
        """
        if x_size <= 0 or y_size <= 0:
            return
        self.worksheet = self.workbook.add_worksheet("Times Tables")
        for x in range(int(x_size)):
            for y in range(int(y_size)):
                color = "#{:02x}{:02x}{:02x}".format(int((x / x_size) * 255),
                                                     int((y / y_size) * 255),
                                                     int(255 - ((x * y) / (x_size * y_size)) * 255))  # type: str
                cell_format = self.workbook.add_format({'bg_color': color})
                self.worksheet.write(y, x, x * y, cell_format)
            self.worksheet.set_column(x, x, len(str(x * y)), None)

    def roll_dice(self, rounds, turns):
        # type: (int, int) -> None
        """
        Simulates two players rolling D20 dice and reports the winner of each turn
        :param rounds: The number of games to be played
        :param turns: The number of turns per game
        :return: None
        """
        if rounds <= 0 or turns <= 0:
            return
        self.worksheet = self.workbook.add_worksheet("Rolling Dice")
        self.worksheet.set_column(0, turns, 3)
        row = 0  # type: int
        for round in range(rounds):  # type: int
            p1_rolls = []  # type: List[int]
            p2_rolls = []  # type: List[int]
            winner = []  # type: List[str]
            for turn in range(turns):
                p1r = random.randint(1, 20)  # type: int
                p2r = random.randint(1, 20)  # type: int
                if p1r == p2r:
                    winner.append("TIE")
                else:
                    winner.append("P1" if p1r > p2r else "P2")
                p1_rolls.append(p1r)
                p2_rolls.append(p2r)
            for data in [p1_rolls, p2_rolls, winner, []]:
                self.worksheet.write_row(row, 0, data)
                row += 1
        format1 = self.workbook.add_format({'bg_color': 'green'})
        format2 = self.workbook.add_format({'bg_color': 'red'})
        format3 = self.workbook.add_format({'bg_color': 'yellow'})
        for i in range(3, 4 * rounds, 4):
            for val, form in [['"P1"', format1],
                              ['"P2"', format2],
                              ['"TIE"', format3]]:
                self.worksheet.conditional_format("{}:{}".format("A" + str(i), self.int_to_column(turns) + str(i)),
                                                  {'type': 'cell',
                                                   'criteria': '==',
                                                   'value': val,
                                                   'format': form})

    def random_tweak(self):
        # type: () -> None
        for sheet in self.workbook.sheetnames:
            s = self.workbook.get_worksheet_by_name(sheet)
            s.write(random.randrange(30), random.randrange(50), 'X')

    @staticmethod
    def int_to_column(num):
        # type: (int) -> str
        """
        Converts an integer to an Excel style column name
        :param num: an integer representing a column number
        :return: A alpha string representing an Excel style column name
        """
        col = ""
        while num >= 702:
            col += string.ascii_uppercase[(num // (26 ** 2)) - 1]
            num %= (26 ** 2)
        while num >= 26:
            col += string.ascii_uppercase[(num // 26) - 1]
            num = num % 26
        col += string.ascii_uppercase[num % 26]
        return col

    @staticmethod
    def column_to_int(column):
        # type: (str) -> None
        """
        Converts an Excel style column string to a numeric value
        :param column: An Excel style column name like "CL" or "D"
        :return: None
        """
        if len(column) < 1 or len(column) > 3:
            return
        column = column.upper()  # in case a lower case column name is submitted
        result = 0
        if len(column) > 2:
            result += 26 ** 2 * (string.ascii_uppercase.index(column[-3]) + 1)
        if len(column) > 1:
            result += 26 * (string.ascii_uppercase.index(column[-2]) + 1)
        return result + string.ascii_uppercase.index(column[-1])

    def itc_test(self):
        # type: () -> None
        """
        A function testing the functionality of the int_to_column algorithm
        :return:
        """
        self.worksheet = self.workbook.add_worksheet("ITC Test")
        for y in range(40):
            for x in range(40):
                self.worksheet.write(y, x, self.int_to_column(x) + str(y + 1))

    def write_words(self):
        # type: () -> None
        """

        :return:
        """
        self.worksheet = self.workbook.add_worksheet("Words")
        infile = open("sortedwords.txt", "r")
        column = []
        all_words = {}
        cur_alpha = 'A'
        width = 0
        longest = ""
        longestwords = []
        for word in infile:
            word = word.strip()
            if word[0].upper() != cur_alpha:
                self.worksheet.write_column(0, self.column_to_int(cur_alpha), column)
                self.worksheet.set_column("{}:{}".format(cur_alpha, cur_alpha), width)
                all_words[cur_alpha] = column
                column = []
                cur_alpha = word[0].upper()
                width = 0
                longestwords.append(longest)
            column.append(word)
            if len(word) > width:
                width = len(word)
                longest = word
        self.worksheet.write_column(0, self.column_to_int("Z"), column)  # for Z column edge case
        self.worksheet.set_column("Z:Z", width)  # for Z column edge case
        longestwords.append(longest)  # for Z column edge case
        self.worksheet.write_column(0, 26, longestwords)


my_book = DemoWorkbook("testFile")
