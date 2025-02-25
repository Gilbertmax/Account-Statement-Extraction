import pdfplumber
import pandas as pd
import re

class PDFExtractor:
    def __init__(self, file_path, bank):
        self.file_path = file_path
        self.bank = bank

    def extract_data(self):
        if self.bank == "BBVA":
            return self.extract_bbva_data()
        else:
            raise ValueError(f"Formato de banco no soportado: {self.bank}")

    def extract_bbva_data(self):
        transactions = []
        with pdfplumber.open(self.file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    transactions.extend(self.process_bbva_text(text))
        return transactions

    def process_bbva_text(self, text):
        lines = text.split("\n")
        extracted_data = []
        current_transaction = []

        for line in lines:
            parts = line.split()
            
            # Detectar transacciones de BBVA (ajusta según el formato real)
            if re.match(r"\d{2}/\d{2}/\d{4}", parts[0]):  # Fecha en formato DD/MM/AAAA
                if current_transaction:
                    extracted_data.append(current_transaction)
                current_transaction = [parts[0]]  # Fecha
                description = " ".join(parts[1:-2])  # Descripción
                amount = parts[-2]  # Monto
                balance = parts[-1]  # Saldo
                current_transaction.extend([description, amount, balance])
            else:
                if current_transaction:
                    current_transaction[1] += " " + " ".join(parts)

        if current_transaction:
            extracted_data.append(current_transaction)

        return extracted_data

    def save_to_excel(self, data, output_path="output.xlsx"):
        if self.bank == "BBVA":
            df = pd.DataFrame(data, columns=["Fecha", "Descripción", "Monto", "Saldo"])
        else:
            df = pd.DataFrame(data)
        df.to_excel(output_path, index=False)
        print(f"Archivo guardado en: {output_path}")

if __name__ == "__main__":
    extractor = PDFExtractor("estado_de_cuenta.pdf", "BBVA")
    data = extractor.extract_data()
    extractor.save_to_excel(data)