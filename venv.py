import re

import xlwings as xw
from PyPDF2 import PdfReader


def main():

    workbook_name = input('Enter Workbook Name:\n')
    sheet_name = input('Enter Worksheet Name:\n')
    scrape_pdf_name = input('Enter PDF Name:\n')

    wb = xw.Book(workbook_name)
    sheet = wb.sheets[sheet_name]
    KPIcount = sheet.range('A1').value #8
    #kpi titles at 8 + KPIcount + 1

    #KeyCount = [0]*KPIcount

    KPIrows = range(7, 7+int(KPIcount))#[7,8,9,10,11,12,13,14]
    print(KPIrows)
    rng = sheet.range('A1:Z100')

    for x in KPIrows:
        KPIcolumns = x-1;
        y1 = int(8 + KPIcount)
        y2 = int(8 + KPIcount)
        print(rng[7,7].value)
        while (True):
            if rng[y2,KPIcolumns].value == None:
                print(rng[y2,KPIcolumns-1].value)
                break
            else:
                #print(rng[y2,x-1].value)
                y2=y2+1
        #y1 is index of the first keyword; y2 is index of the last keyword
        #KeyCount[x-6] = y2-y1
        #for keyword in
            #KeyCount[x-6]:


        keywords_specifier = [KPIcolumns, y1,y2]
        print(keywords_specifier)

        #keywords_spec = sheet.Range((x, y1), (x, y2))

        #if xw.Range('A').xl_range.font_object.color.set(some_rgb_vale)

        #    sheet.range(f'A{search1}')

        Scrape(scrape_pdf_name, keywords_specifier, f'I{x+1}', workbook_name, sheet_name)



class Scrape:
    MAX_DISTANCE = 20  # chars

    def __init__(self, pdf, key_range, destination, workbook_name, sheet_name):
        x = key_range[0]
        y1 = key_range[1]
        y2 = key_range[2]


        #key range consists of two numbers
        self.wb = xw.Book(workbook_name)
        self.sheet = self.wb.sheets[sheet_name]
        rng = self.sheet.range('A1:GZ100')
#        print(rng[17:19, 6:7].value)
        #print(self.sheet.range(x:(x-1), y1:y2).value
        self.keywords = [kwarg for kwarg in rng[y1:y2, x:x+1].value if kwarg]
        #print(self.keywords)
        #print(f"Keywords: {', '.join(self.keywords)}")

        self.pages = [page.extract_text() for page in PdfReader(pdf).pages]
        self.content = '\n'.join(self.pages)
        self.split_content = self.content.split()

        matches = []

        r = re.finditer("[-+]?[.]?[\d]+(?:,\d\d\d)*[\.]?\d*(?:[eE][-+]?\d+)?", self.content) #list of numbers around special units

        for match in r:
            matches.append((match.group(), match.start(0)))
        #for each in r[range(r.__sizeof__())]:
        #print(str(list(r)))
        #print(f'{str(list(r))}')
        r = []
        print(self.keywords)
        #for KEY in self.keywords:
        r.append(re.finditer("|".join(self.keywords), self.content))
        for match in r:
            for h in match:
                matches.append((h.group(), h.start(0)))
            #matches.append((match.group(), match.start(0)))
        matches = sorted(matches, key=lambda x: x[1])


        NumbersFound = []
        for index, match in enumerate(matches):
            if match[0] not in self.keywords:
                continue

            prev_val = matches[index - 1] if matches[index - 1][0] not in self.keywords else (None, float("inf"))
            next_val = matches[index + 1] if matches[index + 1][0] not in self.keywords else (None, float("inf"))


            if (dist := abs(prev_val[1] - match[1])) < abs(next_val[1] - match[1]) and dist < Scrape.MAX_DISTANCE:
                new_entry = f"{prev_val[0]} - {match[0]} "
                print(new_entry)
                NumbersFound.append(new_entry)
            elif abs(prev_val[1] - match[1]) > (dist := abs(next_val[1] - match[1])) and dist < Scrape.MAX_DISTANCE:
                new_entry = f"{match[0]} - {next_val[0]}"
                print(new_entry)
                NumbersFound.append(new_entry)
            print(NumbersFound)
            self.sheet.range(destination).value = NumbersFound
        return




    @staticmethod
    def log_match(key, value):
        print(f"{key} - {value}")

    @staticmethod
    def word_is_value(word):
        return word.replace('.', '', 1).isdigit()



if __name__ == "__main__":
    main()
