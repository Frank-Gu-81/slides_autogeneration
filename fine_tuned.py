from openai import OpenAI

import yaml
from PyPDF2 import PdfReader
import re
from pptx import Presentation
from pptx.util import Inches
import json

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
    max_tokens=4000)
    return response.choices[0].message.content

def save_api_response(response, output_file):
    with open(output_file, 'w') as file:
        file.write(response)

# Main function with modification for system_instruction, task_input, and RAG query
def main(raw_output_file="step_1_structure.txt"):

    with open ("original_report.md", "r") as file:
        md_content = file.read()

    with open("step_1_extracted.txt", "r") as file:
        extracted_structure = file.read()

    # # Step 1: Extract the structure of the markdown content
    # step_1_extract_structure_prompt = step_1_extract_structure(md_content)
    # # Call OpenAI API
    # print("Sending request to OpenAI...")
    # step_1_extracted_structure = call_openai_api(step_1_extract_structure_prompt)
    # save_api_response(step_1_extracted_structure, raw_output_file)

    # Step 2: Outline the structure for PowerPoint slides
    # step_2_outline_slide_structure_prompt = step_2_outline_slide_structure(extracted_structure)
    # Call OpenAI API
    print("Sending request to OpenAI...")
    # step_2_slide_structure = call_openai_api(step_2_outline_slide_structure_prompt)

    step_n_prompt = step_n_generate_slide_json_prompt(extracted_structure)
    step_n_generate_slide_json = call_openai_api(step_n_prompt)
    save_api_response(step_n_generate_slide_json, raw_output_file)




    print("Process completed successfully.")

    
def step_1_extract_structure(md_content):
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

        **Here is the markdown content for you to process and make sure you do not include my prompt in your reseponse:** 
        '''{md_content}'''
        """

    return extract_structure_prompt



def step_n_generate_slide_json_prompt(extracted_structure):
    # Prepare the prompt
    generate_slide_json_prompt = f"""
        **Task:**

        Create a JSON structure for PowerPoint slides based on the provided content. Each slide should have a unique title that best describes the content on that slide, along with a subtitle and at least one bullet point. Ensure content on each slide does not exceed 90 words (excluding the title), and split content logically into multiple slides.

        **Instructions:**

        1. **Structure:**
            - Each slide must have a unique, descriptive title that reflects its specific content.
            - Include at least one subtitle and one bullet point on each slide.
            - Limit content to 90 words per slide (excluding the title).

        2. **Slide Division:**
            - Split content into multiple slides as needed to maintain a smooth and logical flow.
            - Ensure each slide title accurately reflects the specific content it contains.

        3. **Design:**
            - Set title font to 36pt, Dark Blue; subtitle font to 28pt, Black; body text to 20pt, Black.
            - Ensure content fits within the word limit and that slides do not overflow.

        4. **Citations:**
            - Use footnotes ([1], [2], etc.) in the slides and compile all citations on a separate "Citations" slide.

        **Output Format:**
            - Return a valid JSON structure with no additional text. Each slide should include:
                - "Slide Title" (a unique title for each slide)
                - "Formatted Content" (with subtitle and bullet points)
                - "Font Size" and "Text Color" for each element
                - "Overflow Check" (confirm no overflow)
                - "Citations" (if any)

        **Example:**

        [
            {{
                "Slide Title": "AI Solutions Across Industries",
                "Formatted Content": [
                    {{
                        "Subtitle": "Neuron Solutions Overview",
                        "Content": [
                            "Neuron Solutions specializes in AI solutions across industries.",
                            "They work with sectors like pharmaceuticals, energy, and engineering [1]."
                        ],
                        "Font Size": "20pt",
                        "Text Color": "Black"
                    }}
                ],
                "Font Size": "36pt",
                "Text Color": "Dark Blue",
                "Overflow Check": "No overflow",
                "Citations": []
            }},
            {{
                "Slide Title": "Citations",
                "Formatted Content": [
                    {{
                        "Content": [
                            "[1] [Neuron Solutions - Your AI consultant](https://www.neuronsolutions.com)"
                        ],
                        "Font Size": "20pt",
                        "Text Color": "Black"
                    }}
                ],
                "Font Size": "36pt",
                "Text Color": "Dark Blue",
                "Overflow Check": "No overflow"
            }}
        ]

        **Process:**

        Generate the slides in JSON format based on the content, splitting them logically with each slide having a unique, descriptive title, subtitle, and content within 90 words. Return only the JSON output.
        Here is the content:
        '''{extracted_structure}'''
    """


    return generate_slide_json_prompt


def step_2_outline_slide_structure(extracted_structure):
    # Prepare the prompt
    outline_slide_structure_prompt = f"""
        **Task:**

        Map the extracted markdown structure to PowerPoint slides. Create an outline where each item is a slide with its title and content.

        **Instructions:**

        - **H1 Headings:** Each becomes a new slide title.
        - **H2 Headings:** Subsections or new slides if content is extensive.
        - **H3+ Headings:** Bullet points or sub-bullets.
        - **Paragraphs:** Include as slide content without summarizing.
        - **Lists:** Convert into bullet points, maintaining hierarchy.
        - **Formatting:** Preserve bold and italic text.
        - **Links:** Include display text; note URLs if relevant.
        - **Images:** Indicate inclusion with alt text and URL.
        - **Code Blocks:** Include if essential.
        - **Tables:** Convert into slide content.

        **Guidelines:**

        - Do **not** condense content.
        - Avoid overcrowding; split slides if needed.
        - Maintain logical flow from the markdown.

        **Output Format:**

        - Numbered list; each number is a slide.
        - For each slide:
        - **Slide Title:** from H1 heading
        - **Content:** include elements per instructions
        - **Notes:** for additional info such as reference URLs

        **Process:**

        Use these instructions to map the extracted markdown structure to slides. Do not return this prompt.

        Here is the extracted structure:

        '''{extracted_structure}'''
    """

    return outline_slide_structure_prompt




if __name__ == "__main__":
      # Replace with your OpenAI API key
    # Provide the path to the YAML and PDF files

    # Call the main function to process the inputs and generate the YAML output
    main(raw_output_file="revised_slides.json")
