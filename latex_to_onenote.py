import sys
import os
import tempfile
import subprocess
import pyperclip
import win32clipboard
from docx import Document
from lxml import etree

def latex_to_onenote_clipboard(latex_string):
    """
    Convert LaTeX equation to OMML format and place on clipboard for OneNote
    """
    # Clean string (remove $ delimiters if present)
    latex_clean = latex_string.strip().strip('$').strip()
    
    # Create temp dir and md file for pandoc conversion
    with tempfile.TemporaryDirectory() as tmpdir:
        md_file = os.path.join(tmpdir, 'input.md')
        docx_file = os.path.join(tmpdir, 'output.docx')
        
        # Write in inline math format
        with open(md_file, 'w', encoding='utf-8') as f:
            f.write(f'${latex_clean}$')
        
        # Convert with pandoc to docx (which contains OMML)
        pandoc_path = r"C:\Users\Justin\AppData\Local\Pandoc\pandoc.exe"
        if not os.path.exists(pandoc_path):
             # Fallback to system path
             pandoc_path = 'pandoc'

        try:
            result = subprocess.run(
                [pandoc_path, md_file, '-o', docx_file],
                capture_output=True,
                text=True,
                check=True
            )
        except subprocess.CalledProcessError as e:
            print(f"Pandoc error: {e.stderr}")
            return False
        except FileNotFoundError:
            print("Error: Pandoc not found. Please install Pandoc first.")
            return False
        
        # Open docx and extract OMML
        try:
            doc = Document(docx_file)
            
            # Get underlying XML
            # OMML equations are stored in paragraph._element
            omml_found = False
            for paragraph in doc.paragraphs:
                # Look for math elements in the paragraph
                para_xml = paragraph._element
                
                # Find oMath elements (OMML namespace)
                ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
                math_elements = para_xml.findall('.//m:oMath', ns)
                
                if math_elements:
                    # Extract the OMML
                    math_element = math_elements[0]
                    
                    # Convert to string and fix namespace for OneNote (OMML 2004 vs OpenXML 2006)
                    raw_omml = etree.tostring(math_element, encoding='unicode')
                    raw_omml = raw_omml.replace(
                        'http://schemas.openxmlformats.org/officeDocument/2006/math',
                        'http://schemas.microsoft.com/office/2004/12/omml'
                    )
                    
                    # Wrap in oMathPara for OneNote
                    omml_string = f'<m:oMathPara xmlns:m="http://schemas.microsoft.com/office/2004/12/omml">{raw_omml}</m:oMathPara>'
                    
                    # Wrap in HTML format for clipboard
                    html_output = f'<!--[if gte msEquation 12]>{omml_string}<![endif]-->'
                    
                    # Set HTML clipboard
                    set_html_clipboard(html_output)
                    omml_found = True
                    print("✓ Converted LaTeX to OMML and copied to clipboard!")
                    print(f"Original LaTeX: {latex_clean}")
                    print("Now paste into OneNote with Ctrl+V")
                    return True
            
            if not omml_found:
                print("Warning: No equation found in converted document")
                return False
                
        except Exception as e:
            print(f"Error extracting OMML: {e}")
            return False

def set_html_clipboard(html_string):
    """
    Set HTML format on clipboard (for OneNote to recognize OMML)
    """
    # HTML clipboard format requires specific header
    html_header = "Version:0.9\r\nStartHTML:0000000000\r\nEndHTML:0000000000\r\nStartFragment:0000000000\r\nEndFragment:0000000000\r\n"
    html_prefix = "<html><body><!--StartFragment-->"
    html_suffix = "<!--EndFragment--></body></html>"
    
    full_html = html_header + html_prefix + html_string + html_suffix
    
    # Calculate byte positions
    html_bytes = full_html.encode('utf-8')
    start_html = html_header.index('StartHTML:') + 10
    end_html = html_header.index('EndHTML:') + 8
    start_frag = html_header.index('StartFragment:') + 14
    end_frag = html_header.index('EndFragment:') + 12
    
    # Update with actual positions
    full_html = full_html.replace('0000000000', str(len(html_header)).zfill(10), 1)
    full_html = full_html.replace('0000000000', str(len(html_bytes)).zfill(10), 1)
    full_html = full_html.replace('0000000000', str(len(html_header + html_prefix)).zfill(10), 1)
    full_html = full_html.replace('0000000000', str(len(html_header + html_prefix + html_string)).zfill(10), 1)
    
    # Set clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
    win32clipboard.SetClipboardData(cf_html, full_html.encode('utf-8'))
    win32clipboard.CloseClipboard()

if __name__ == "__main__":
    # Get LaTeX from clipboard or command line argument
    if len(sys.argv) > 1:
        latex_input = sys.argv[1]
    else:
        latex_input = pyperclip.paste()
        print(f"Reading from clipboard: {latex_input}")
    
    if not latex_input.strip():
        print("Error: No LaTeX input provided")
        print("Usage: python latex_to_onenote.py 'V_{rms}=\\frac{V_{peak}}{\\sqrt{2}}'")
        print("   or: Copy LaTeX to clipboard and run: python latex_to_onenote.py")
        sys.exit(1)
    
    success = latex_to_onenote_clipboard(latex_input)
    sys.exit(0 if success else 1)
