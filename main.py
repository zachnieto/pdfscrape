import re

import xlwings as xw
from PyPDF2 import PdfReader


class Scrape:
    MAX_DISTANCE = 20  # chars

    def __init__(self, pdf):
        self.wb = xw.Book(r"keywords.xlsx")
        self.sheet = self.wb.sheets[0]
        self.keywords = [kwarg for kwarg in self.sheet.range("A:A")[1:500].value if kwarg]

        print(f"Keywords: {', '.join(self.keywords)}")

        self.pages = [page.extract_text() for page in PdfReader(pdf).pages]
        self.content = '\n'.join(self.pages)
        self.split_content = self.content.split()

        matches = []

        r = re.finditer("[-+]?[.]?[\d]+(?:,\d\d\d)*[\.]?\d*(?:[eE][-+]?\d+)?", self.content)
        for match in r:
            matches.append((match.group(), match.start(0)))

        r = re.finditer("|".join(self.keywords), self.content)
        for match in r:
            matches.append((match.group(), match.start(0)))

        matches = sorted(matches, key=lambda x: x[1])

        for index, match in enumerate(matches):
            if match[0] not in self.keywords:
                continue

            prev_val = matches[index - 1] if matches[index - 1][0] not in self.keywords else (None, float("inf"))
            next_val = matches[index + 1] if matches[index + 1][0] not in self.keywords else (None, float("inf"))

            if (dist := abs(prev_val[1] - match[1])) < abs(next_val[1] - match[1]) and dist < Scrape.MAX_DISTANCE:
                print(f"{match[0]} - {prev_val[0]}")
            elif abs(prev_val[1] - match[1]) > (dist := abs(next_val[1] - match[1])) and dist < Scrape.MAX_DISTANCE:
                print(f"{match[0]} - {next_val[0]}")

    @staticmethod
    def log_match(key, value):
        print(f"{key} - {value}")

    @staticmethod
    def word_is_value(word):
        return word.replace('.', '', 1).isdigit()


if __name__ == "__main__":
    Scrape("term-loan-credit-agreement.pdf")
