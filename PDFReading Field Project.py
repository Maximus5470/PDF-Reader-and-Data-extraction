import pdfplumber, pandas as pd, csv, re
import tkinter as tk
from tkinter import filedialog, messagebox


def select_pdfs():
    sources = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if not sources:
        messagebox.showerror("Error", "No PDFs selected.")
        exit()

    output_file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not output_file:
        messagebox.showerror("Error", "No output file selected.")
        exit()
    return sources, output_file


def main(source):
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

        # Extract the data from tables
        tables = pdf.pages[0].extract_tables()

        columns = [i.lower() for i in tables[0][0]]
        df = pd.DataFrame(tables[0][1:], columns=columns)
        if "no" in columns or "sno" in columns or "s.no" in columns:
            df.drop("no", axis=1, inplace=True)
        df.index = df.index + 1

        last_row = list(df.loc[len(tables[0]) - 1])  # extracting total
        df.drop(len(tables[0]) - 1, inplace=True)

        total_text = ""
        for i in last_row:
            if (i != "") and (i is not None):
                total_text += i
        return Name, InvoiceNumber, total_text, df


def CSVWrite(output_file, Name, InvoiceNumber, total_text, df):
    with open(output_file, "a", newline="") as file:
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
        writer.writerow({})


root = tk.Tk()

root.title("PDF to CSV Converter")
select_button = tk.Button(root, text="Select PDFs and Convert to CSV", command=select_pdfs)
select_button.pack(pady=20)

sources, output_file = select_pdfs()
for i in sources:
    Name, InvoiceNumber, total_text, df = main(i)
    CSVWrite(output_file, Name, InvoiceNumber, total_text, df)
