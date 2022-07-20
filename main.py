from PyPDF2 import PdfReader


class Scrape:

    def __init__(self, pdf):
        self.keywords = open("keywords.txt", "r").read().split("\n")

        reader = PdfReader(pdf)
        self.pages = [page.extract_text().split() for page in reader.pages]
        self.content = [word for page in self.pages for word in page]
        # self.content = "Test is the value 4. 7 matches to Hello. Word is not with Test but with 1. Test. 3 is Dog".split()

        matching = None

        for word in self.content:
            word = word.replace('.', '', 1).replace(",", '').replace("$", "")
            if word in self.keywords:
                if matching and self.word_is_value(matching):
                    self.log_match(word, matching)
                    matching = None
                else:
                    matching = word
            elif self.word_is_value(word):
                if matching and matching in self.keywords:
                    self.log_match(matching, word)
                    matching = None
                else:
                    matching = word

            if matching and matching[-1] == ".":
                matching = None

    @staticmethod
    def log_match(key, value):
        print(f"{key} - {value}")

    @staticmethod
    def word_is_value(word):
        return word.replace('.', '', 1).isdigit()


if __name__ == "__main__":
    Scrape("test.pdf")
