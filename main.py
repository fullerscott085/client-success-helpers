import fitz  # PyMuPDF
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Iterator
from enum import Enum
import pandas as pd
from dataclasses import asdict
import zipfile
import re
from io import BytesIO


class DataType(Enum):
    """Enum representing valid data types for extracted values."""
    TEXT = 'text'
    NUMBER = 'number'
    DATE = 'date'


@dataclass
class KeyItem:
    """Represents a single searchable item to extract from documents."""
    key: str
    search_text: str
    data_type: DataType = DataType.TEXT
    result: Optional[str] = field(default=None, repr=False)
    file_name: Optional[str] = field(default=None, repr=False)
    
    def __post_init__(self):
        """Validate key format on initialization."""
        if not self.key.islower() or ' ' in self.key:
            raise ValueError("Key must be lowercase without spaces")

@dataclass
class KeyItemCollection:
    """Manages a collection of KeyItems with validation and access methods."""
    items: List[KeyItem] = field(default_factory=list)
    
    def __getitem__(self, key: str) -> KeyItem:
        """Get item by key."""
        for item in self.items:
            if item.key == key:
                return item
        raise KeyError(f"No item with key '{key}' found")
    
    def add_item(self, item: KeyItem):
        """Add a new item with unique key validation."""
        if any(existing.key == item.key for existing in self.items):
            raise ValueError(f"Key '{item.key}' already exists")
        self.items.append(item)

    def search_texts(self) -> List[str]:
        """Get all search texts from the collection."""
        return [item.search_text for item in self.items]

    def update_result(self, key: str, value: str):
        """Update the result for a specific key."""
        self[key].result = value


def parse_invoice_text(matched_text):
    """Parse the invoice text into a structured dictionary."""
    # Clean and split the text
    lines = [line.strip() for line in matched_text.split('\n') if line.strip()]
    
    # Find where values begin (after "Line Total")
    try:
        value_start = lines.index('Line Total') + 1
    except ValueError:
        return None
    
    headers = lines[:value_start]
    values = lines[value_start:]
    
    # Define field structure with expected headers
    field_definitions = [
        {'name': 'Line', 'headers': ['Line']},
        {'name': 'Marketing Part Number', 'headers': ['Marketing', 'Part', 'Number']},
        {'name': 'Description', 'headers': ['Description'], 'multi_line': True},
        {'name': 'Comprising Manufacturing Part No', 'headers': ['Comprising', 'Manufacturing', 'Part', 'No']},
        {'name': 'Assembled In', 'headers': ['Assembled', 'In']},
        {'name': 'Qty Per Country', 'headers': ['Qty', 'Per', 'Country']},
        {'name': 'Shipped Qty', 'headers': ['Shipped', 'Qty']},
        {'name': 'Unit Price', 'headers': ['Unit', 'Price']},
        {'name': 'Line Total', 'headers': ['Line', 'Total']}
    ]
    
    result = {}
    value_index = 0
    
    for field in field_definitions:
        if value_index >= len(values):
            result[field['name']] = None
            continue
        
        # Handle multi-line fields (like Description)
        if field.get('multi_line'):
            description_parts = []
            while value_index < len(values):
                # Stop when we hit a numeric value (next field)
                if (values[value_index].replace(',', '').replace('.', '').isdigit() or
                    values[value_index].startswith('CN')):  # Special case for country code
                    break
                description_parts.append(values[value_index])
                value_index += 1
            result[field['name']] = ' '.join(description_parts)
        else:
            # Single value fields
            result[field['name']] = values[value_index]
            value_index += 1
    
    return result

def extract_text_from_pdf(doc_pages):
    texts = []
    for doc_page in doc_pages:
        texts.append(doc_page.get_text())

    return texts


def process_key_value_pairs(text_pages: List[str], collection: KeyItemCollection, 
                          search_mapping: Dict[str, str]) -> None:
    """Extract key-value pairs from document text pages."""
    key_found = None
    for page_text in text_pages:
        for line in page_text.split('\n'):
            clean_line = line.strip()
            
            if key_found:
                collection.update_result(
                    key=search_mapping[key_found],
                    value=clean_line
                )
                del search_mapping[key_found]
                key_found = None
                continue
                
            if clean_line in search_mapping:
                key_found = clean_line


def extract_invoice_body(text_pages: List[str]) -> Optional[Dict]:
    """Extract structured invoice data from text pages."""
    for page_text in text_pages:
        match = re.search(r'(Line.*?)License:', page_text, re.DOTALL)
        if match:
            return parse_invoice_text(match.group(1))
    return None


def process_single_pdf(file: BytesIO, collection: KeyItemCollection) -> Dict:
    """Process a single PDF file and return extracted data."""
    results = {}
    
    try:
        with fitz.open(stream=file) as document:
            text_pages = extract_text_from_pdf(document)
            
            # Create fresh mapping for this document
            search_mapping = {item.search_text: item.key for item in collection.items}
            
            # Process key-value pairs
            process_key_value_pairs(text_pages, collection, search_mapping)
            
            # Extract structured invoice data
            results.update(extract_invoice_body(text_pages) or {})
            
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
    
    return results


def process_zip_archive(zip_path: str, collection: KeyItemCollection, progress_callback=None) -> List[Dict]:
    """Process all PDFs in a ZIP archive and return extracted data.
    
    Args:
        zip_path: Path to ZIP file containing PDFs
        collection: Configured KeyItemCollection with search items
        
    Returns:
        List of dictionaries containing extracted data from all files
    """
    results = []
    
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        pdf_files = [file_info for file_info in zip_ref.infolist() if file_info.filename.lower().endswith('.pdf')]
        total_files = len(pdf_files)

        for idx, file_info in enumerate(zip_ref.infolist(), start=1):
            if not file_info.filename.lower().endswith('.pdf'):
                continue
                
            print(f" Processing {file_info.filename} ".center(80, '-'))
            
            with zip_ref.open(file_info.filename) as file:
                # Process the PDF and get results
                pdf_data = BytesIO(file.read())
                invoice_data = process_single_pdf(pdf_data, collection)
                
                # Combine all results
                document_results = {
                    'filename': file_info.filename,
                    **{item.key: item.result for item in collection.items},
                    **invoice_data
                }
                results.append(document_results)
                
                # Reset collection results for next document
                for item in collection.items:
                    item.result = None

                # Update progress in Streamlit
                if progress_callback:
                    progress_callback(idx, total_files)
                    
    return results


def dataframes_for_export(results: list):
    df = pd.DataFrame(results)
    df[['Unit Price', 'Line Total']] = (
        df[['Unit Price', 'Line Total']]
        .apply(lambda x: pd.to_numeric(x.str.replace(',', ''), errors='coerce'))
    )
    df['oracle-so#'] = df['comm-inv-no'].apply(lambda x: '-'.join(str(x).strip().split('-')[1:]))

    mapping_from_pdf_to_salesforce_excel_upload = {
            'Comprising Manufacturing Part No': 'PN', 
            'Description': 'Desc', 
            'Shipped Qty': 'QTY', 
            'Unit Price': 'Value', 
            'Line Total': 'Total', 
            'oracle-so#': 'Oracle SO#'
    }
    
    df_for_salesforce = df[list(mapping_from_pdf_to_salesforce_excel_upload.keys())].rename(
        mapping_from_pdf_to_salesforce_excel_upload, axis=1
        )
    
    return df_for_salesforce, df
