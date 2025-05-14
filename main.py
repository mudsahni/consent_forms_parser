import os
import json
import argparse
import tempfile
from typing import Dict, List, Optional
import base64
from google import genai
from google.genai import types, Client
import pikepdf
import json
import argparse
from typing import Dict, List, Any
import json
import csv
import sys

# Dictionary of MIME types
MIME_TYPES = {
    'pdf': 'application/pdf',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'png': 'image/png',
    'tiff': 'image/tiff',
    'tif': 'image/tiff',
    'bmp': 'image/bmp',
    'gif': 'image/gif',
    'webp': 'image/webp',
    'doc': 'application/msword',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
}


class DocumentProcessor:
    def __init__(self, api_key: str, model_name: str = "gemini-2.0-flash"):
        self.api_key = api_key
        self.client = genai.Client(api_key=api_key)
        self.model_name = model_name
        # Using tempfile.TemporaryDirectory for robust cleanup
        self._temp_dir_obj = tempfile.TemporaryDirectory()
        self.temp_dir = self._temp_dir_obj.name
        print(f"Created temporary directory for intermediate operations: {self.temp_dir}")

    def __del__(self):
        """Clean up temporary files on object destruction"""
        try:
            self._temp_dir_obj.cleanup()
            print(f"Removed temporary directory: {self.temp_dir}")
        except Exception as e:
            print(f"Error cleaning up temp directory {self.temp_dir}: {str(e)}")

    def _get_mime_type(self, file_name: str) -> Optional[str]:
        """
        Determine MIME type based on file extension
        """
        extension = file_name.lower().split('.')[-1]
        mime_type = MIME_TYPES.get(extension)

        if not mime_type:
            print(f"Unsupported file format: {extension}")

        return mime_type

    def convert_docx_to_pdf(self, docx_path: str) -> Optional[str]:
        """
        Convert a Word document (.docx or .doc) to PDF format, saving it in the same directory.
        Returns path to the created PDF or None if conversion failed.
        .doc files are significantly harder to convert reliably without tools like MS Office or LibreOffice.
        This function will primarily focus on .docx and attempt .doc with LibreOffice if available.
        """
        original_filename_lower = os.path.basename(docx_path).lower()
        base_name = os.path.basename(docx_path).rsplit('.', 1)[0]
        output_directory = os.path.dirname(docx_path)
        # Ensure output_directory is absolute if docx_path is relative and CWD changes
        if not output_directory:  # If docx_path is just "file.docx"
            output_directory = "."
        pdf_path = os.path.join(output_directory, f"{base_name}.pdf")

        print(f"Target PDF path for {docx_path} is {pdf_path}")

        # Method 1: Using docx2pdf library (primarily for .docx)
        if original_filename_lower.endswith('.docx'):
            try:
                from docx2pdf import convert
                print(f"Attempting to convert {docx_path} to {pdf_path} using docx2pdf...")
                convert(docx_path, pdf_path)
                if os.path.exists(pdf_path):
                    print(f"Successfully converted {docx_path} to {pdf_path} using docx2pdf.")
                    return pdf_path
                else:
                    print(f"docx2pdf conversion ran for {docx_path} but output PDF not found at {pdf_path}.")
            except ImportError:
                print("docx2pdf library not found. Skipping docx2pdf conversion for .docx.")
            except Exception as e:
                print(f"Error converting .docx to PDF using docx2pdf for {docx_path}: {str(e)}")
        elif original_filename_lower.endswith('.doc'):
            print(f".doc file detected ({docx_path}). docx2pdf does not directly support .doc. Will try LibreOffice.")

        # Method 2: Alternative conversion using LibreOffice (if installed) - works for .doc and .docx
        try:
            import subprocess
            print(f"Attempting to convert {docx_path} using LibreOffice, outputting to {output_directory}...")
            result = subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', output_directory, docx_path
            ], capture_output=True, text=True, timeout=120)  # Increased timeout

            if result.returncode == 0:
                # LibreOffice should produce a file with the same base name and .pdf extension
                if os.path.exists(pdf_path):
                    print(f"Successfully converted {docx_path} to {pdf_path} using LibreOffice.")
                    return pdf_path
                else:
                    # Sometimes LibreOffice might slightly alter the name (e.g. spaces)
                    # A more robust check might be needed if base_name can have complex chars
                    found_pdf = None
                    for f_name in os.listdir(output_directory):
                        if f_name.lower().startswith(base_name.lower()) and f_name.lower().endswith(".pdf"):
                            generated_pdf_path_check = os.path.join(output_directory, f_name)
                            # Basic check to see if it's likely our file (e.g. by modification time or size)
                            # For now, just take the first match.
                            if os.path.exists(generated_pdf_path_check):
                                print(f"LibreOffice converted {docx_path}, found as {generated_pdf_path_check}")
                                return generated_pdf_path_check  # Return the actual name if different
                    if not found_pdf:
                        print(
                            f"LibreOffice conversion reported success for {docx_path}, but expected PDF {pdf_path} (or similar) not found. Stdout: {result.stdout} Stderr: {result.stderr}")
            else:
                print(
                    f"LibreOffice conversion failed for {docx_path}. Return code: {result.returncode}. Stderr: {result.stderr}")
        except FileNotFoundError:
            print("LibreOffice not found or not in PATH. Skipping LibreOffice conversion.")
        except subprocess.TimeoutExpired:
            print(f"LibreOffice conversion for {docx_path} timed out.")
        except Exception as e:
            print(f"Error converting DOC(X) to PDF using LibreOffice for {docx_path}: {str(e)}")

        # Method 3: Extract text and create a simple PDF with reportlab (text only, for .docx)
        if original_filename_lower.endswith('.docx'):
            try:
                from docx import Document  # For text extraction from .docx
                from reportlab.pdfgen import canvas
                from reportlab.lib.pagesizes import letter
                from reportlab.lib.units import inch

                print(f"Attempting to create a text-only PDF for {docx_path} at {pdf_path} using ReportLab...")
                doc = Document(docx_path)
                text_content = "\n".join([para.text for para in doc.paragraphs])

                c = canvas.Canvas(pdf_path, pagesize=letter)
                text_object = c.beginText(0.75 * inch, 10.25 * inch)  # Margins
                text_object.setFont("Helvetica", 10)

                lines = text_content.split('\n')
                for line in lines:
                    # Handle potential non-ASCII characters by attempting to encode and decode robustly
                    try:
                        line_to_write = line.encode('utf-8', 'replace').decode('utf-8', 'ignore')
                    except Exception:  # Broad exception for encoding issues
                        line_to_write = "".join(char if ord(char) < 128 else '?' for char in line)

                    # Basic line wrapping (can be improved with reportlab's Paragraph for better flow)
                    max_width = letter[0] - 1.5 * inch  # Available width

                    current_segment = ""
                    for word in line_to_write.split(' '):
                        if not current_segment:
                            test_segment = word
                        else:
                            test_segment = current_segment + " " + word

                        if c.stringWidth(test_segment, "Helvetica", 10) <= max_width:
                            current_segment = test_segment
                        else:
                            text_object.textLine(current_segment)
                            current_segment = word  # Start new line with current word
                            if text_object.getY() < 0.75 * inch:
                                c.drawText(text_object)
                                c.showPage()
                                text_object = c.beginText(0.75 * inch, 10.25 * inch)
                                text_object.setFont("Helvetica", 10)

                    if current_segment:  # Add any remaining part of the line
                        text_object.textLine(current_segment)

                    if text_object.getY() < 0.75 * inch and line_to_write:  # Ensure Y check makes sense
                        c.drawText(text_object)
                        c.showPage()
                        text_object = c.beginText(0.75 * inch, 10.25 * inch)
                        text_object.setFont("Helvetica", 10)

                c.drawText(text_object)
                c.save()
                if os.path.exists(pdf_path):
                    print(f"Successfully created simple text-only PDF {pdf_path} from {docx_path} using ReportLab.")
                    return pdf_path
                else:
                    print(f"ReportLab conversion ran for {docx_path} but output PDF not found at {pdf_path}.")

            except ImportError:
                print("python-docx or reportlab library not found. Skipping ReportLab PDF creation for .docx.")
            except Exception as e:
                print(f"Error creating simple PDF with ReportLab from {docx_path}: {str(e)}")
        elif original_filename_lower.endswith('.doc'):
            print(f"ReportLab fallback for .doc is not implemented (requires .docx for text extraction).")

        print(f"All PDF conversion methods failed for {docx_path}.")
        return None

    def is_pdf_digitally_signed(self, pdf_path: str) -> bool:
        """
        Check if a PDF file contains digital signatures.
        Returns True if digitally signed, False otherwise.
        """
        try:
            with pikepdf.open(pdf_path) as pdf:
                # Check for /ByteRange and /Contents in signature dictionaries
                for page in pdf.pages:
                    for annot in page.get('/Annots', []):
                        annot_obj = annot.get_object()
                        if annot_obj.get('/Subtype') == '/Widget' and annot_obj.get('/FT') == '/Sig':
                            if '/ByteRange' in annot_obj and '/Contents' in annot_obj:
                                return True

                # Check signature dictionaries in document catalog
                if '/AcroForm' in pdf.Root:
                    if '/Fields' in pdf.Root.AcroForm:
                        for field in pdf.Root.AcroForm.Fields:
                            field_obj = field.get_object()
                            if field_obj.get('/FT') == '/Sig':
                                if '/ByteRange' in field_obj and '/Contents' in field_obj:
                                    return True

                # Check for signature dictionary in document catalog
                if '/Perms' in pdf.Root:
                    perms = pdf.Root.Perms
                    if isinstance(perms, pikepdf.Dictionary):
                        if '/DocMDP' in perms:
                            return True

            return False

        except Exception as e:
            print(f"Error checking for digital signatures in {pdf_path}: {str(e)}")
            return False

    def _validate_json_response(self, response_text: str, file_name: str) -> str:
        """
        Validate that the response is a valid JSON string
        """
        try:
            # Attempt to parse the string as JSON
            json.loads(response_text)
            return response_text
        except json.JSONDecodeError as e:
            error_msg = f"Invalid JSON response for file {file_name}: {str(e)}"
            print(error_msg)
            print(f"Response content: {response_text[:200]}...")  # Log first 200 chars
            raise Exception(error_msg)


    def process_file(self, file_path: str) -> Dict:
        """
        Process a file using Gemini
        """
        file_name = os.path.basename(file_path)

        if file_name.lower().endswith('.docx') or file_name.lower().endswith('.doc'):
            # Convert DOCX to PDF
            pdf_path = self.convert_docx_to_pdf(file_path)
            if pdf_path:
                file_path = pdf_path
                file_name = os.path.basename(pdf_path)
            else:
                print(f"Failed to convert {file_name} to PDF.")
                return {
                    "file_path": file_path,
                    "error": "Failed to convert DOCX to PDF",
                    "has_verification": False,
                    "has_declaration": False,
                    "verification_signature_present": False,
                    "declaration_signature_present": False,
                    "digitally_signed": False
                }
        # Check if PDF is digitally signed
        is_signed = False
        if file_name.lower().endswith('.pdf'):
            is_signed = self.is_pdf_digitally_signed(file_path)
            if is_signed:
                print(f"PDF {file_name} is digitally signed. No need to check with Gemini.")
                return {
                    "file_path": file_path,
                    "has_verification": True,
                    "has_declaration": True,  # Can't determine without Gemini
                    "verification_signature_present": True,
                    "declaration_signature_present": False,  # Can't determine without Gemini
                    "digitally_signed": True,
                    "gemini_response": None
                }

        # Get MIME type
        mime_type = self._get_mime_type(file_name)
        if not mime_type:
            return {
                "file_path": file_path,
                "error": f"Unsupported file format: {file_name.split('.')[-1]}",
                "has_verification": False,
                "digitally_signed": is_signed
            }

        # Read file content
        with open(file_path, 'rb') as f:
            file_content = f.read()

        # Prepare the prompt
        prompt = """
        Analyze this document to identify verification/declaration sections and signatures. Focus on:

        1. TITLES: Locate any sections explicitly titled 'Verification', 'Declaration', 'Attestation', 'Certification', 'Affidavit', 'Sworn Statement', or similar legal confirmation headings.

        2. SIGNATURES: Identify any:
           - Actual handwritten signatures (not typed names)
           - Digital signatures
           - Signature blocks with completed signatures
           - Signature images

        3. RELATIONSHIP: When finding verification/declaration sections, check subsequent pages for associated signatures.

        4. CLAIM INFORMATION: Extract any monetary values labeled as 'claim amount', 'claimed sum', 'requested amount', 'total claim', or similar terminology.

        5. CONFIDENCE ASSESSMENT: Evaluate your confidence in each finding on a scale from 0.0 to 1.0, where:
           - 0.0-0.3: Low confidence (unclear or ambiguous elements)
           - 0.4-0.7: Medium confidence (somewhat clear but with potential uncertainty)
           - 0.8-1.0: High confidence (clearly identified elements)

        Return ONLY a valid JSON object with this exact structure:
        {
          "has_verification_section": boolean,
          "has_declaration_section": boolean,
          "verification_signature_present": boolean,
          "declaration_signature_present": boolean,
          "claim_amount": number or null,
          "claim_currency": string or null,
          "confidence_score": number (0.0-1.0)
        }

        Provide ONLY the JSON object with no additional explanations, comments, or formatting.
        """
        try:
            # Send to Gemini
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=[
                    types.Part.from_bytes(
                        data=file_content,
                        mime_type=mime_type,
                    ),
                    prompt
                ]
            )

            # Extract and clean the response text
            response_text = response.text

            if not isinstance(response_text, str):
                print(f"Error parsing file {file_name}: Response is not a string")
                return {
                    "file_path": file_path,
                    "error": "Response is not a string",
                    "has_verification": False,
                    "has_declaration": False,
                    "verification_signature_present": False,
                    "declaration_signature_present": False,
                    "digitally_signed": is_signed,
                    "claim_amount": None
                }

            # Clean up the response if it contains Markdown code blocks
            json_response = str(response_text).replace("```json", "").replace("```", "").strip()

            # Validate and parse JSON
            validated_response = self._validate_json_response(json_response, file_name)
            gemini_result = json.loads(validated_response)

            # Combine our results with Gemini's response
            return {
                "file_path": file_path,
                "has_verification": gemini_result.get("has_verification", False),
                "has_declaration": gemini_result.get("has_declaration", False),
                "verification_signature_present": gemini_result.get("verification_signature_present", False),
                "declaration_signature_present": gemini_result.get("declaration_signature_present", False),
                "digitally_signed": is_signed,
                "claim_amount": gemini_result.get("claim_amount", None),
                "gemini_response": gemini_result
            }

        except Exception as e:
            print(f"Error processing {file_name}: {str(e)}")
            return {
                "file_path": file_path,
                "error": str(e),
                "has_verification": False,
                "has_declaration": False,
                "verification_signature_present": False,
                "declaration_signature_present": False,
                "digitally_signed": is_signed,
                "claim_amount": None
            }

    def process_folder(self, folder_path: str) -> List[Dict]:
        """
        Process all files in a folder and its subfolders that have 'formd' in their name
        """
        results = []

        # Validate folder exists
        if not os.path.isdir(folder_path):
            print(f"Folder not found: {folder_path}")
            return results

        # Walk through the directory tree to process all files including those in subfolders
        for root, dirs, files in os.walk(folder_path):
            for filename in files:
                # Check if 'formd' is in the filename (case-insensitive)
                if 'formd' in filename.lower().replace("_","").replace("-", "").replace(" ", ""):
                    file_path = os.path.join(root, filename)

                    # Print relative path for better readability
                    rel_path = os.path.relpath(file_path, folder_path)
                    print(f"Processing: {rel_path}")

                    # Process the file
                    result = self.process_file(file_path)
                    results.append(result)

                    # Optional: Print quick result
                    if result.get('error'):
                        print(f"  Error: {result['error']}")
                    else:
                        verification = "✓" if result.get('has_verification') else "✗"
                        declaration = "✓" if result.get('has_declaration') else "✗"
                        print(f"  Verification: {verification}, Declaration: {declaration}")

        # If no files were found, inform the user
        if not results:
            print(f"No files with 'formd' in their name found in {folder_path} or its subfolders")

        return results

def validate_and_reformat(input_data: Dict[str, List[Dict[str, Any]]]) -> Dict[str, List[Dict[str, Any]]]:
    """
    Validate and reformat the input JSON data according to specified rules.

    Rules:
    1. If gemini_response.has_verification_title && gemini_response.verification_signature_present
       then verification_signature_present = true, else false
    2. Same for declaration

    Final JSON should include:
    - digitally_signed
    - verification_signature_present
    - declaration_signature_present
    - confidence
    - file_path
    """
    results = []

    for item in input_data.get("results", []):
        file_path = item.get("file_path", "")
        digitally_signed = item.get("digitally_signed", False)
        claim_amount = item.get("claim_amount", None)
        # Get Gemini response
        gemini_response = item.get("gemini_response", {})

        if gemini_response:
            # Apply validation rules exactly as specified
            verification_signature_present = bool(
                gemini_response.get("has_verification_title", False) and
                gemini_response.get("verification_signature_present", False)
            )

            declaration_signature_present = bool(
                gemini_response.get("has_declaration_title", False) and
                gemini_response.get("declaration_signature_present", False)
            )

            # Get confidence score from Gemini response
            confidence = gemini_response.get("confidence", 0.0)
        else:
            # If no Gemini response (e.g., digitally signed PDF)
            verification_signature_present = False
            declaration_signature_present = False

            # If digitally signed, we'll set confidence to 1.0
            confidence = 1.0 if digitally_signed else 0.0

        # Create the reformatted result with only the required fields
        reformatted_result = {
            "file_path": file_path,
            "digitally_signed": digitally_signed,
            "verification_signature_present": verification_signature_present,
            "declaration_signature_present": declaration_signature_present,
            "confidence": confidence,
            "claim_amount": claim_amount,
        }

        results.append(reformatted_result)

    return {"results": results}


def json_to_csv(json_file_path: str, csv_file_path: str):
    """
    Convert a JSON file to a CSV file.

    Args:
        json_file_path (str): Path to the input JSON file
        csv_file_path (str): Path to the output CSV file
    """
    try:
        # Read JSON data
        with open(json_file_path, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
            data = data["results"] if "results" in data else data

        # Ensure data is a list of objects
        if not isinstance(data, list):
            # If it's a single object, wrap it in a list
            if isinstance(data, dict):
                data = [data]
            # If it's something else, try to handle common structures
            elif isinstance(data, dict) and any(isinstance(data[key], list) for key in data):
                # Find the first list in the dictionary
                for key in data:
                    if isinstance(data[key], list):
                        data = data[key]
                        break
            else:
                raise ValueError("JSON data structure not supported. Please provide an array of objects.")

        # If data is empty, return
        if not data:
            print("The JSON data is empty.")
            return

        # Extract field names from the first object
        field_names = list(data[0].keys())

        # Write to CSV
        with open(csv_file_path, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=field_names)
            writer.writeheader()
            writer.writerows(data)

        print(f"Successfully converted {json_file_path} to {csv_file_path}")

    except FileNotFoundError:
        print(f"Error: File {json_file_path} not found.")
    except json.JSONDecodeError:
        print(f"Error: {json_file_path} is not a valid JSON file.")
    except Exception as e:
        print(f"Error: {str(e)}")


def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Process documents with Gemini')
    parser.add_argument('--folder', type=str, required=True, help='Path to folder containing files')
    parser.add_argument('--api-key', type=str, required=True, help='Gemini API key')
    parser.add_argument('--model', type=str, default='gemini-2.0-flash', help='Gemini model name')
    parser.add_argument('--output', type=str, default='results.json', help='Output JSON file path')
    parser.add_argument('--validated-output', type=str, default='validated.json', help='Validated output JSON file path')

    args = parser.parse_args()
    api_key_to_use = args.api_key if args.api_key else os.environ.get("GOOGLE_API_KEY")
    if not api_key_to_use:
        print("Error: Gemini API key is required. Provide it with --api-key or set the GOOGLE_API_KEY environment variable.")
        return 1

    try:
        # Initialize processor
        processor = DocumentProcessor(api_key=args.api_key, model_name=args.model)

        # Process the folder
        results = processor.process_folder(args.folder)
        # Save results to file
        with open(args.output, 'w') as f:
            json.dump({"results": results}, f, indent=2)

        print(f"Processed {len(results)} files. Results saved to {args.output}")

        # Validate and reformat
        validated_data = validate_and_reformat({"results": results})
        with open(args.validated_output, 'w') as f:
            json.dump(validated_data, f, indent=2)
        print(f"Validation complete. Results saved to {args.output}")

        # Convert validated JSON to CSV
        csv_file_path = args.validated_output.replace('.json', '.csv')
        json_to_csv(args.validated_output, csv_file_path)
        print(f"Converted validated JSON to CSV: {csv_file_path}")

        # Print summary
        verification_count = sum(1 for r in results if r.get('has_verification', False))
        declaration_count = sum(1 for r in results if r.get('has_declaration', False))
        verification_signed_count = sum(1 for r in results if r.get('verification_signature_present', False))
        declaration_signed_count = sum(1 for r in results if r.get('declaration_signature_present', False))
        error_count = sum(1 for r in results if 'error' in r)

        print(f"Summary:")
        print(f"- {verification_count} files have verification sections")
        print(f"- {declaration_count} files have declaration sections")
        print(f"- {verification_signed_count} files have signed verification sections")
        print(f"- {declaration_signed_count} files have signed declaration sections")
        print(f"- {error_count} errors encountered")
    except Exception as e:
        print(f"Error: {str(e)}")
        return 1
    finally:
        return 0

if __name__ == "__main__":
    main()