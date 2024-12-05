import os
import json
import pandas as pd
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
import matplotlib.pyplot as plt

# Azure Form Recognizer credentials
endpoint = "https://digestibilitydataextraction.cognitiveservices.azure.com/"  # Azure endpoint
key = "6G7oGuGkFSyqTHneA9bZ4NkdWccliEak813aLNt0Lm8Z1mijG6BaJQQJ99AKACBsN54XJ3w3AAALACOG3n9P"  # Azure key
model_id = "NewmodelTable"

# Initialize the DocumentAnalysisClient
document_analysis_client = DocumentAnalysisClient(
    endpoint=endpoint, credential=AzureKeyCredential(key)
)

# Input and output folder paths
pdf_folder_path = "./PDF"  # Folder containing input PDFs
json_folder_path = "./JSON"  # Folder to save JSON files
output_excel_path = "./final_result.xlsx"  # Final Excel file path
calculate_path = "./R/final_result.xlsx"

# Ensure the output folder exists
os.makedirs(json_folder_path, exist_ok=True)

def analyze_pdf_and_save_json(pdf_path, json_path):
    """Analyzes a PDF using Azure Form Recognizer and saves the result as a JSON file."""
    with open(pdf_path, "rb") as pdf_file:
        poller = document_analysis_client.begin_analyze_document(
            model_id=model_id, document=pdf_file
        )
        result = poller.result()

    # Save the JSON response
    with open(json_path, "w") as json_file:
        json.dump(result.to_dict(), json_file, indent=4)
    print(f"JSON saved to {json_path}.")


def process_json_to_side_by_side_excel(json_folder_path, output_file):
    """Processes JSON files to extract tables and saves them side-by-side in an Excel file."""
    required_headers = [
        "ASP", "THR", "SER", "GLU", "PRO", "GLY", "ALA", "CYS",
        "VAL", "MET", "ILE", "LEU", "TYR", "PHE", "HIS", "LYS",
        "ARG", "TRP"
    ]
    alternative_headers = [
        "Asp", "Thr", "Ser", "Glu", "Pro", "Gly", "Ala", "Cys",
        "Val", "Met", "Ile", "Leu", "Try", "Phe", "His", "Lys",
        "Arg", "Trp"
    ]
    all_possible_headers = set(required_headers + alternative_headers)
    skip_terms = ["article in", "abstract", "article info"]

    filtered_tables = []

    # Iterate through JSON files
    for file_name in os.listdir(json_folder_path):
        if file_name.endswith(".json"):
            file_path = os.path.join(json_folder_path, file_name)

            # Load the JSON file
            with open(file_path, "r") as file:
                data = json.load(file)

            # Extract tables from the JSON
            tables = data.get("tables", [])

            for table in tables:
                rows = {}
                for cell in table.get("cells", []):
                    row_index = cell.get("row_index")
                    column_index = cell.get("column_index")
                    content = cell.get("content", "")

                    if row_index not in rows:
                        rows[row_index] = {}

                    rows[row_index][column_index] = content

                df = pd.DataFrame.from_dict(rows, orient="index").sort_index(axis=1)

                if df.iloc[0, 0].strip().lower() in skip_terms:
                    continue

                has_pdcaas = any(
                    df.iloc[0].str.contains("PDCAAS", case=False, na=False)
                ) or df.apply(
                    lambda x: x.str.contains("PDCAAS", case=False, na=False).any(), axis=0
                ).any()

                headers_top = df.iloc[0].str.upper().isin(all_possible_headers).sum() == len(required_headers)
                headers_bottom = df.iloc[-1].str.upper().isin(all_possible_headers).sum() == len(required_headers)

                if headers_top:
                    df.columns = df.iloc[0]
                    df = df[1:]
                elif headers_bottom:
                    df.columns = df.iloc[-1]
                    df = df[:-1]

                if headers_top or headers_bottom or has_pdcaas:
                    filtered_tables.append(df)

    if filtered_tables:
        max_rows = max([df.shape[0] for df in filtered_tables])
        aligned_tables = [
            df.reset_index(drop=True).reindex(range(max_rows), fill_value="")
            for df in filtered_tables
        ]

        combined_df = pd.concat(aligned_tables, axis=1)
        combined_df.to_excel(output_file, index=False)
        print(f"Filtered tables saved to {output_file}.")
    else:
        print("No relevant tables found.")


def calculate_and_update_excel(file_path):
    """Calculate PDCAAS and IVPDCAAS, and generate graphs."""
    df = pd.read_excel(file_path)

    if "ASS" in df.columns and "TPD" in df.columns:
        df["PDCAAS"] = (df["ASS"] * df["TPD"]) / 100
    else:
        df["PDCAAS"] = None

    if "ASS" in df.columns and "IVPD" in df.columns:
        df["IVPDCAAS"] = (df["ASS"] * df["IVPD"]) / 100
    else:
        df["IVPDCAAS"] = None

    df.to_excel(file_path, index=False)
    print(f"Updated Excel file saved with PDCAAS and IVPDCAAS: {file_path}")

    #create_graphs(df, os.path.dirname(file_path))


def create_graphs(df, output_folder):
    """Generate graphs from the Excel data."""
    if "SAMPLE" in df.columns and "PDCAAS" in df.columns:
        plt.figure()
        df.groupby("SAMPLE")["PDCAAS"].mean().plot(kind="bar", title="PDCAAS by SAMPLE")
        plt.xlabel("Sample")
        plt.ylabel("PDCAAS")
        plt.tight_layout()
        plt.savefig(os.path.join(output_folder, "PDCAAS_by_SAMPLE.png"))
        print("Graph saved: PDCAAS by SAMPLE")

    if "TPD" in df.columns and "PDCAAS" in df.columns:
        plt.figure()
        plt.scatter(df["TPD"], df["PDCAAS"])
        plt.title("PDCAAS by TPD")
        plt.xlabel("TPD")
        plt.ylabel("PDCAAS")
        plt.tight_layout()
        plt.savefig(os.path.join(output_folder, "PDCAAS_by_TPD.png"))
        print("Graph saved: PDCAAS by TPD")

    if "ASS" in df.columns and "PDCAAS" in df.columns:
        plt.figure()
        plt.scatter(df["ASS"], df["PDCAAS"])
        plt.title("ASS by PDCAAS")
        plt.xlabel("ASS")
        plt.ylabel("PDCAAS")
        plt.tight_layout()
        plt.savefig(os.path.join(output_folder, "ASS_by_PDCAAS.png"))
        print("Graph saved: ASS by PDCAAS")


# Process each PDF in the folder
for pdf_file in os.listdir(pdf_folder_path):
    if pdf_file.endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder_path, pdf_file)
        json_path = os.path.join(json_folder_path, f"{os.path.splitext(pdf_file)[0]}.json")
        analyze_pdf_and_save_json(pdf_path, json_path)

# Process all JSON files and save the final Excel file
process_json_to_side_by_side_excel(json_folder_path, output_excel_path)

# Perform calculations and update Excel with graphs
calculate_and_update_excel(calculate_path)
