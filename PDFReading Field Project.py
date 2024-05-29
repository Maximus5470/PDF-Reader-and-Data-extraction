import pdfplumber, pandas as pd, csv, re

sources = [
    "sample1.pdf",
    "sample2.pdf",
    "sample3.pdf"
]
for source in sources:
    with pdfplumber.open(source) as pdf:
        # Extract the text
        page = pdf.pages[0].extract_text()
        text = page.split("\n")
        Name = None
        InvoiceNumber = None

        for num, line in enumerate(text):
            if "invoice to" in line.lower() or "billed to" in line.lower() or "bill to" in line.lower():
                if re.search(r"^\s+$", text[num + 1]):
                    Name = " ".join(text[num + 1].strip().split()[:-1])
                    break
                elif re.search(r"^#\d+", text[num + 1]):
                    InvoiceNumber = text[num + 1].strip().split()[-1][1:]
                    break
                else:
                    Name = " ".join(text[num + 1].strip().split()[:-1])
                    InvoiceNumber = text[num + 1].strip().split()[-1][1:]
                    break

            elif "mr" in line.lower() or "ms" in line.lower() or "mrs" in line.lower() or "dr" in line.lower():
                Name = line.strip()
                InvoiceNumber = text[num + 1].strip().split()[-1][1:]
                break

        print(f"Name: {Name}")
        print(f"Invoice Number: {InvoiceNumber}")
        # Extract the data from tables
        tables = pdf.pages[0].extract_tables()

        columns = [i.lower() for i in tables[0][0]]
        df = pd.DataFrame(tables[0][1:], columns=columns)
        if "no" in columns or "sno" in columns or "s.no" in columns:
            df.drop("no", axis=1, inplace=True)
        df.index = df.index + 1

        last_row = list(df.loc[len(tables[0]) - 1])  #extracting total
        df.drop(len(tables[0]) - 1, inplace=True)

        print(df)
        total_text = ""
        for i in last_row:
            if (i != "") and (i is not None):
                total_text += i

        print(total_text)
        with open("C://Users/GAUTHAM SHARMA/Documents/ReportCSV.csv", "a", newline="") as file:
            dictlist = [{
                "Name": Name,
                "Invoice Number": InvoiceNumber,
                "Total": total_text
            }]
            fields = ["Name", "Invoice Number", "Total"]

            writer = csv.DictWriter(file, fieldnames=fields)
            writer.writeheader()
            writer.writerows(dictlist)

            df.to_csv(file, index=False)
