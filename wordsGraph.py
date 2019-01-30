import xlsxwriter  # type: ignore
import matplotlib.pyplot as plt  # type: ignore
from typing import Optional
import string
import numpy


class DemoWorkbook(object):
    def __init__(self, filename):
        # type: (str) -> None
        self.workbook = xlsxwriter.Workbook(filename + ".xlsx", None)
        self.worksheet = None
        self.write_words()
        self.workbook.close()

    def write_words(self):
        # type: () -> None
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
        all_words[cur_alpha] = column  # for Z edge case
        self.worksheet.set_column("Z:Z", width)  # for Z column edge case
        longestwords.append(longest)  # for Z column edge case
        self.worksheet.write_column(0, 26, longestwords)

        labels = all_words.keys()
        bars = [len(x) for x in all_words.values()]
        print labels
        print bars
        print zip(labels, bars)
        plt.bar(numpy.arange(26), bars)
        plt.xticks(numpy.arange(26), labels)
        plt.title('Words Per Letter In English')
        plt.show()

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
        # type: (str) -> Optional[int]
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


my_book = DemoWorkbook("wordsList")
