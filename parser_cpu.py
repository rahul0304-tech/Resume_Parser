import spacy
from spacy.tokens import Span
from spacy.util import filter_spans
import pandas as pd
import PyPDF2
from docxtpl import DocxTemplate
import os

@spacy.Language.component("merge_entities")
def merge_entities(doc):
    spans = list(doc.ents) + list(doc.noun_chunks)
    spans = filter_spans(spans)
    
    with doc.retokenize() as retokenizer:
        for span in spans:
            retokenizer.merge(span)
    
    return doc

class EntityGenerator:
    __slots__ = ['text']
    
    def __init__(self, text=None):
        self.text = text
        
    def get(self):
        """
        Return a dictionary of entities and their values
        """
        nlp = spacy.load("en_core_web_trf")  # Load the transformer-based model
        nlp.add_pipe("merge_entities", last=True)  # Add the merge_entities component

        doc = nlp(self.text)

        entities = {}
        for ent in doc.ents:
            if ent.label_ not in entities:
                entities[ent.label_] = []
            entities[ent.label_].append(ent.text.strip())

        return entities

class Resume:
    def __init__(self, filename=None):
        self.filename = filename
        
    def get(self):
        """
        Extract text from the resume file.
        """
        if self.filename.lower().endswith('.pdf'):
            return self._extract_text_from_pdf()
        elif self.filename.lower().endswith('.docx'):
            return self._extract_text_from_docx()
        elif self.filename.lower().endswith('.txt'):
            return self._extract_text_from_txt()
        else:
            raise ValueError("Unsupported file format")

    def _extract_text_from_pdf(self):
        with open(self.filename, 'rb') as fFileObj:
            pdfReader = PyPDF2.PdfReader(fFileObj)
            if len(pdfReader.pages) == 0:
                raise ValueError("The PDF file is empty.")
            pageObj = pdfReader.pages[0]
            print(f"Total Pages: {len(pdfReader.pages)}")
            resume = pageObj.extract_text()
        return resume
    
    def _extract_text_from_docx(self):
        doc = DocxTemplate(self.filename)
        full_text = []
        for para in doc.docx.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)

    def _extract_text_from_txt(self):
        with open(self.filename, 'r', encoding='utf-8') as file:
            return file.read()

def json_to_excel(json_data, filenames, excel_filename):
    # Flatten the JSON data
    flattened_data = {}
    ordered_entities = ["PERSON", "ORG", "GPE", "PRODUCT", "DATE"]
    
    for idx, data in enumerate(json_data):
        resume_key = os.path.basename(filenames[idx].strip())  # Use the filename as the key
        flattened_data[resume_key] = {}
        for entity, values in data.items():
            if entity in ordered_entities:
                flattened_data[resume_key][entity] = values

    # Create a DataFrame to store the data in a structured way
    df_list = []

    for resume_key, entities in flattened_data.items():
        df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in entities.items()]))
        df_list.append(df)

    # Concatenate all dataframes
    final_df = pd.concat(df_list, ignore_index=True)
    
    # Reorder the columns to match the ordered_entities list
    final_df = final_df.reindex(columns=ordered_entities, fill_value='')

    # Write DataFrame to Excel
    final_df.to_excel(excel_filename, index=False)

def main():
    filenames = input("Enter the paths of the resume files (PDF, DOCX, TXT) separated by commas: ").split(',')
    excel_filename = input("Enter the name of the Excel file: ").strip()
    
    all_responses = []

    for filename in filenames:
        filename = filename.strip()  # Remove any leading/trailing whitespace
        if not os.path.exists(filename):
            print(f"Error: File '{filename}' does not exist. Skipping.")
            continue

        resume = Resume(filename=filename)
        try:
            response_news = resume.get()
        except Exception as e:
            print(f"Error processing file '{filename}': {e}")
            continue

        helper = EntityGenerator(text=response_news)
        response = helper.get()
        all_responses.append(response)

    if all_responses:
        # Convert combined JSON response to Excel
        json_to_excel(all_responses, filenames, excel_filename)
        print(f"JSON data has been successfully written to {excel_filename}")
    else:
        print("No valid resumes were processed. Excel file was not created.")

if __name__ == "__main__":
    main()
