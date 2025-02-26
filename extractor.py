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
        elif self.bank == "Banco del Bajío":
            return self.extract_banco_bajio_data()
        elif self.bank == "Scotiabank":
            return self.extract_scotiabank_data()
        elif self.bank == "BBASE":
            return self.extract_bbase_data()
        elif self.bank == "Banorte":
            return self.extract_banorte_data()
        else:
            raise ValueError(f"Formato de banco no soportado: {self.bank}")

    def extract_bbva_data(self):
        data = {"Información General": {}, "Transacciones": [], "Resumen Financiero": {}}
        with pdfplumber.open(self.file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    self.process_bbva_text(text, data)
        return data

    def process_bbva_text(self, text, data):
        lines = text.split("\n")
        for line in lines:
            # Extraer información general
            if "No. de Cuenta" in line and ":" in line:
                data["Información General"]["Número de Cuenta"] = line.split(":")[1].strip()
            if "No. de Cliente" in line and ":" in line:
                data["Información General"]["Número de Cliente"] = line.split(":")[1].strip()
            if "R.F.C" in line and ":" in line:
                data["Información General"]["RFC"] = line.split(":")[1].strip()

            # Extraer transacciones
            if re.match(r"\d{2}/\w{3}", line):  # Fecha en formato DD/MMM
                parts = line.split()
                if len(parts) >= 4:
                    fecha = parts[0]
                    descripcion = " ".join(parts[1:-2])
                    cargos = parts[-2] if parts[-2].replace(",", "").replace(".", "").isdigit() else "0"
                    abonos = parts[-1] if parts[-1].replace(",", "").replace(".", "").isdigit() else "0"
                    data["Transacciones"].append([fecha, descripcion, cargos, abonos])

            # Extraer resumen financiero
            if "Saldo Anterior" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Anterior"] = line.split(":")[1].strip()
            if "Saldo Final" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Final"] = line.split(":")[1].strip()

    def extract_banco_bajio_data(self):
        data = {"Información General": {}, "Transacciones": [], "Resumen Financiero": {}}
        with pdfplumber.open(self.file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    self.process_banco_bajio_text(text, data)
        return data

    def process_banco_bajio_text(self, text, data):
        lines = text.split("\n")
        for line in lines:
            # Extraer información general
            if "NÚMERO DE CLIENTE" in line and ":" in line:
                data["Información General"]["Número de Cliente"] = line.split(":")[1].strip()
            if "R.F.C." in line and ":" in line:
                data["Información General"]["RFC"] = line.split(":")[1].strip()

            # Extraer transacciones
            if re.match(r"\d{2} \w{3}", line):  # Fecha en formato DD MMM
                parts = line.split()
                if len(parts) >= 5:
                    fecha = parts[0] + " " + parts[1]  # Fecha en formato DD MMM
                    descripcion = " ".join(parts[2:-2])
                    depositos = parts[-2] if parts[-2].replace(",", "").replace(".", "").isdigit() else "0"
                    retiros = parts[-1] if parts[-1].replace(",", "").replace(".", "").isdigit() else "0"
                    data["Transacciones"].append([fecha, descripcion, depositos, retiros])

            # Extraer resumen financiero
            if "SALDO ANTERIOR" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Anterior"] = line.split(":")[1].strip()
            if "SALDO ACTUAL" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Actual"] = line.split(":")[1].strip()

    def extract_scotiabank_data(self):
        data = {"Información General": {}, "Transacciones": [], "Resumen Financiero": {}}
        with pdfplumber.open(self.file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    self.process_scotiabank_text(text, data)
        return data

    def process_scotiabank_text(self, text, data):
        lines = text.split("\n")
        for line in lines:
            # Extraer información general
            if "No. de Cuenta" in line and ":" in line:
                data["Información General"]["Número de Cuenta"] = line.split(":")[1].strip()
            if "RFC Cliente" in line and ":" in line:
                data["Información General"]["RFC"] = line.split(":")[1].strip()

            # Extraer transacciones
            if re.match(r"\d{2}-\w{3}-\d{2}", line):  # Fecha en formato DD-MMM-AA
                parts = line.split()
                if len(parts) >= 4:
                    fecha = parts[0]
                    descripcion = " ".join(parts[1:-2])
                    cargos = parts[-2] if parts[-2].replace(",", "").replace(".", "").isdigit() else "0"
                    abonos = parts[-1] if parts[-1].replace(",", "").replace(".", "").isdigit() else "0"
                    data["Transacciones"].append([fecha, descripcion, cargos, abonos])

            # Extraer resumen financiero
            if "Saldo inicial" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Inicial"] = line.split(":")[1].strip()
            if "Saldo final" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Final"] = line.split(":")[1].strip()

    def extract_bbase_data(self):
        data = {"Información General": {}, "Transacciones": [], "Resumen Financiero": {}}
        with pdfplumber.open(self.file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    self.process_bbase_text(text, data)
        return data

    def process_bbase_text(self, text, data):
        lines = text.split("\n")
        for line in lines:
            # Extraer información general
            if "Número de Cuenta" in line and ":" in line:
                data["Información General"]["Número de Cuenta"] = line.split(":")[1].strip()
            if "RFC" in line and ":" in line:
                data["Información General"]["RFC"] = line.split(":")[1].strip()

            # Extraer transacciones
            if re.match(r"\d{2}-\w{3}-\d{2}", line):  # Fecha en formato DD-MMM-AA
                parts = line.split()
                if len(parts) >= 5:
                    fecha = parts[0]
                    concepto = " ".join(parts[1:-3])
                    cargo = parts[-3] if parts[-3].replace(",", "").replace(".", "").isdigit() else "0"
                    abonos = parts[-2] if parts[-2].replace(",", "").replace(".", "").isdigit() else "0"
                    saldo = parts[-1] if parts[-1].replace(",", "").replace(".", "").isdigit() else "0"
                    data["Transacciones"].append([fecha, concepto, cargo, abonos, saldo])

            # Extraer resumen financiero
            if "Saldo Anterior" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Anterior"] = line.split(":")[1].strip()
            if "Saldo Actual" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Actual"] = line.split(":")[1].strip()

    def extract_banorte_data(self):
        data = {"Información General": {}, "Transacciones": [], "Resumen Financiero": {}}
        with pdfplumber.open(self.file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    self.process_banorte_text(text, data)
        return data

    def process_banorte_text(self, text, data):
        lines = text.split("\n")
        for line in lines:
            # Extraer información general
            if "Número de Cuenta" in line and ":" in line:
                data["Información General"]["Número de Cuenta"] = line.split(":")[1].strip()
            if "RFC" in line and ":" in line:
                data["Información General"]["RFC"] = line.split(":")[1].strip()

            # Extraer transacciones
            if re.match(r"\d{2}-\w{3}-\d{2}", line):  # Fecha en formato DD-MMM-AA
                parts = line.split()
                if len(parts) >= 5:
                    fecha = parts[0]
                    descripcion = " ".join(parts[1:-3])
                    monto_deposito = parts[-3] if parts[-3].replace(",", "").replace(".", "").isdigit() else "0"
                    monto_retiro = parts[-2] if parts[-2].replace(",", "").replace(".", "").isdigit() else "0"
                    saldo = parts[-1] if parts[-1].replace(",", "").replace(".", "").isdigit() else "0"
                    data["Transacciones"].append([fecha, descripcion, monto_deposito, monto_retiro, saldo])

            # Extraer resumen financiero
            if "Saldo Anterior" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Anterior"] = line.split(":")[1].strip()
            if "Saldo Actual" in line and ":" in line:
                data["Resumen Financiero"]["Saldo Actual"] = line.split(":")[1].strip()

    def save_to_excel(self, data, output_path="output.xlsx"):
        with pd.ExcelWriter(output_path) as writer:
            # Guardar información general
            if "Información General" in data:
                df_info = pd.DataFrame(list(data["Información General"].items()), columns=["Campo", "Valor"])
                df_info.to_excel(writer, sheet_name="Información General", index=False)

            # Guardar transacciones
            if "Transacciones" in data:
                if self.bank == "BBVA":
                    df_trans = pd.DataFrame(data["Transacciones"], columns=["Fecha", "Descripción", "Cargos", "Abonos"])
                elif self.bank == "Banco del Bajío":
                    df_trans = pd.DataFrame(data["Transacciones"], columns=["Fecha", "Descripción", "Depósitos", "Retiros"])
                elif self.bank == "Scotiabank":
                    df_trans = pd.DataFrame(data["Transacciones"], columns=["Fecha", "Descripción", "Cargos", "Abonos"])
                elif self.bank == "BBASE":
                    df_trans = pd.DataFrame(data["Transacciones"], columns=["Fecha", "Concepto", "Cargo", "Abonos", "Saldo"])
                elif self.bank == "Banorte":
                    df_trans = pd.DataFrame(data["Transacciones"], columns=["Fecha", "Descripción/Establecimiento", "Monto del Depósito", "Monto del Retiro", "Saldo"])
                df_trans.to_excel(writer, sheet_name="Transacciones", index=False)

            # Guardar resumen financiero
            if "Resumen Financiero" in data:
                df_resumen = pd.DataFrame(list(data["Resumen Financiero"].items()), columns=["Campo", "Valor"])
                df_resumen.to_excel(writer, sheet_name="Resumen Financiero", index=False)

        print(f"Archivo guardado en: {output_path}")

if __name__ == "__main__":
    extractor = PDFExtractor("estado_de_cuenta.pdf", "BBASE")
    data = extractor.extract_data()
    extractor.save_to_excel(data)