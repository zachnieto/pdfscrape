import os
import re

import xlwings as xw
from PyPDF2 import PdfReader


class Scrape:
    MAX_DISTANCE = 50  # chars

    def __init__(self, pdf):
        self.wb = xw.Book(r"keywords.xlsm")
        self.sheet = self.wb.sheets[0]
        self.keywords = [kwarg for kwarg in self.sheet.range("H:H")[7:15].value if kwarg]

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


if __name__ == "__main__":
    while 1:
        pdf = input("PDF filename: ")
        if os.path.exists(pdf):
            break
        print("File not found, please try again")

    Scrape(pdf)
    input("Press enter to exit...")
