from openai import OpenAI

import yaml
from PyPDF2 import PdfReader
import re
from pptx import Presentation
from pptx.util import Inches
import json
import re

with open('config.yaml', 'r') as file:
    config = yaml.safe_load(file)
client = OpenAI(api_key=config["openai_key"])

# Function to send a prompt to OpenAI API and get a response
def call_openai_api(prompt):
    response = client.chat.completions.create(model="gpt-4",
    messages=[
        {"role": "system", "content": "You are an assistant that converts markdown content into a detailed, structured, and visually appealing powerpoint presentation."},
        {"role": "user", "content": prompt}
    ],
    max_tokens=4000,
    temperature=0.2)
    return response.choices[0].message.content

def save_api_response(response, output_file):
    with open(output_file, 'w') as file:
        file.write(response)

def acquire_sections(markdown_file):
    sections = re.split(r'(^#+ .*$)', markdown_file, flags=re.MULTILINE)
    return sections

def step_1_extract_structure(md_content, title="Extract Markdown Structure"):
    # Prepare the prompt
    extract_structure_prompt = f"""
        **Your Task:**

        Please parse the following markdown content and extract its complete hierarchical structure, including:

        - **Headings:** All levels from H1 to H6.
        - **Lists:** Bullet points and numbered lists, including nested lists.
        - **Images:** Include alt text and image URLs.
        - **Code Blocks:** Both inline code and fenced code blocks.
        - **Blockquotes:** Any quoted text.
        - **Tables:** Include table headers and cell content.
        - **Emphasis:** Bold, italics, and other text formatting.
        - **Links:** Hyperlinks with their display text and URLs.
        - **Paragraphs:** Regular text content.

        **Instructions:**

        1. **Organize Hierarchically:**
        - Use indentation to represent the hierarchy of the content.
        - Reflect the nesting of headings and subheadings.
        - Maintain the order of elements as they appear in the markdown.

        2. **Element Identification:**
        - **Headings:** Indicate the level (e.g., H1, H2, H3) and the heading text.
        - **Lists:** Differentiate between bullet points (`-`, `*`, `+`) and numbered lists (`1.`, `2.`).
            - For nested lists, increase the indentation.
        - **Images:** Provide the alt text and the image URL.
        - **Code Blocks:**
            - **Inline Code:** Enclosed in single backticks (`).
            - **Fenced Code Blocks:** Enclosed in triple backticks (```).
            - Include the language identifier if specified (e.g., ```python).
        - **Blockquotes:** Indicate quoted text with a clear marker.
        - **Tables:** Represent the structure of tables, including headers and cells.
        - **Emphasis:** Note any text formatting like bold or italics.
        - **Links:** Provide the display text and the URL.
        - **Paragraphs:** Include the full text content.

        3. **Formatting Guidelines:**
        - **Indentation:** Use consistent indentation (e.g., two or four spaces) for each level.
        - **Bullets and Numbers:** Use bullet points (`-`) for unordered lists and numbers (`1.`, `2.`) for ordered lists.
        - **Clarity:** Clearly label each element type for easy identification.
        - **Completeness:** Do not omit any content; ensure all information from the markdown is included.

        4. **Example Format:**

        - **H1:** Introduction
            - **Paragraph:** This section introduces the main topics.
            - **H2:** Background
            - **Paragraph:** Provides background information.
            - **Bullet List:**
                - Point one
                - Point two
                - Sub-point one
                - Sub-point two
            - **Image:** 
                - **Alt Text:** Diagram of process
                - **URL:** http://example.com/image.png
            - **Code Block (Language: Python):**
                ```python
                def example_function():
                    pass
                ```
            - **Blockquote:**
            - "This is a quoted text."

        5. **Output Presentation:**

        - Present the structure in an indented outline format.
        - Use labels to identify element types.
        - Ensure readability and clear hierarchy.

        **Here is the title and markdown content for you to process and make sure you do not include my prompt in your reseponse:** 
        
        '''{title + md_content}'''
        """

    return extract_structure_prompt

def append_response_to_file(api_response, output_file):
    with open(output_file, 'a') as file:
        file.write(api_response)
        file.write("\n")

def iterative_structure_extraction(markdown_content):
    sections = acquire_sections(markdown_content)
    for i in range(1, len(section_content), 2):
        print(f"Processing section {i}...")
        section_title = sections[i].strip()
        section_content = sections[i+1].strip()
        prompt = step_1_extract_structure(title=section_title, md_content=section_content)
        response = call_openai_api(prompt)
        print("current response: ", response)
        append_response_to_file(response, 'extracted_structure.txt')
        print("=====================================")

def iterative_ppt_generation(structured_file):
    pass

if __name__ == "__main__":
    # Load the markdown file
    with open('original_report.md', 'r') as file:
        markdown_content = file.read()

    # Perform iterative structure extraction
    iterative_structure_extraction(markdown_content)
    

    