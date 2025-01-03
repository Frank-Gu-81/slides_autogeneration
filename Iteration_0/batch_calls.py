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

def step_n_generate_slide_json_prompt(extracted_structure):

    generate_slide_json_prompt = f"""
            **Task:**

            Create a JSON structure for PowerPoint slides based on the provided content. Your goal is to transform the content into slides that include all key information without unnecessary summarization. Each slide should have a unique title reflecting its specific content, along with subtitles and bullet points that capture all important details. Use hierarchical bullet points to represent any nested information.

            **Instructions:**

            1. **Content Inclusion:**
            - **Include all key points, facts, and details** from the content.
            - **Avoid unnecessary summarization**; ensure important information is retained.
            - Use bullet points to break down the content into digestible pieces.
            - **Preserve quotes, citations, and specific terminology** as they appear in the content.

            2. **Slide Structure:**
            - Each slide should have:
                - A unique, descriptive **Slide Title**.
                - **Formatted Content** containing:
                - **Subtitles** where appropriate.
                - **Bullet Points** with hierarchical structure to represent nested information.
            - Use hierarchical bullet points for any lists or detailed explanations.

            3. **Design Specifications:**
            - **Slide Title**: Font Size 36pt, Dark Blue.
            - **Subtitles**: Font Size 28pt, Black.
            - **Bullet Points**: Font Size 20pt, Black.
            - **Text Color**: As specified above.

            4. **Slide Division:**
            - Split the content logically into multiple slides to ensure clarity.
            - **If a slide becomes too content-heavy, split it into multiple slides** while maintaining logical flow.
            - **Do not exceed 7 bullet points per slide** to avoid overcrowding.

            5. **Citations and Links:**
            - Include citations and hyperlinks as footnotes ([1], [2], etc.) in the bullet points.
            - Compile all citations on a separate "Citations" slide at the end.

            6. **Formatting Details:**
            - **Maintain the original wording** as much as possible.
            - **Retain the structure** of the content, including any emphasis or lists.

            **Output Format:**

            - Return **only** a valid JSON array with no additional text or explanations.
            - Each slide should be a JSON object containing:
            - `"Slide Title"`: The title of the slide.
            - `"Formatted Content"`: A list of content sections, each with:
                - `"Subtitle"`: (optional) Subtitle text.
                - `"Content"`: A list of bullet points (strings), including hierarchical bullet points represented as nested lists.
                - `"Font Size"`: Font size for the content (e.g., "20pt").
                - `"Text Color"`: Text color (e.g., "Black").
            - `"Font Size"`: Font size for the slide title (e.g., "36pt").
            - `"Text Color"`: Text color for the slide title (e.g., "Dark Blue").
            - `"Overflow Check"`: Confirm "No overflow" or indicate if content needs to be split further.
            - `"Citations"`: List of citations used on the slide.

            **Example:**

            [
                {{
                    "Slide Title": "Executive Summary",
                    "Formatted Content": [
                        {{
                            "Subtitle": "Neuron Solutions Overview",
                            "Content": [
                                "Neuron Solutions, operating under the brand neuron.ai, is a consulting firm specializing in AI solutions across various industries.",
                                [
                                    "Provides services from conceptualization to implementation of AI projects.",
                                    "Enables businesses to develop their own AI capabilities or integrate supplier AI systems into operations ([1]).",
                                    "Client sectors include pharmaceuticals, energy, and engineering."
                                ],
                                "Successfully enhanced operational efficiencies through AI-driven solutions."
                            ],
                            "Font Size": "20pt",
                            "Text Color": "Black"
                        }}
                    ],
                    "Font Size": "36pt",
                    "Text Color": "Dark Blue",
                    "Overflow Check": "No overflow",
                    "Citations": ["[1]"]
                }},
                {{
                    "Slide Title": "Innovation in Finance",
                    "Formatted Content": [
                        {{
                            "Subtitle": "AI Transforming Finance",
                            "Content": [
                                "Leverages AI to:",
                                [
                                    "Transform data analysis.",
                                    "Improve investment accessibility.",
                                    "Redefine banking operations."
                                ],
                                "Participates in industry conferences like Future of Finance 2024 ([2])."
                            ],
                            "Font Size": "20pt",
                            "Text Color": "Black"
                        }},
                        {{
                            "Subtitle": "AI Launchpad Methodology",
                            "Content": [
                                "Combines human intelligence with AI capabilities.",
                                "Optimizes work processes and enhances customer relations.",
                                "Positions Neuron Solutions as a key player in AI consulting."
                            ],
                            "Font Size": "20pt",
                            "Text Color": "Black"
                        }}
                    ],
                    "Font Size": "36pt",
                    "Text Color": "Dark Blue",
                    "Overflow Check": "No overflow",
                    "Citations": ["[2]"]
                }},
                {{
                    "Slide Title": "Citations",
                    "Formatted Content": [
                        {{
                            "Content": [
                                "[1] [Neuron Solutions - Your AI consultant](https://www.neuronsolutions.com)",
                                "[2] [Artificial Intelligence and the Future of Finance - Neuron Solutions](https://www.neuronsolutions.com/finance)"
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

            Transform the provided content into slides as per the instructions, ensuring that **all key information is included and properly structured**. **Do not omit important details or excessively summarize the content**. Use hierarchical bullet points to represent nested information. **Maintain the original wording where possible**. **Return only the JSON output and nothing else.**

            Here is the content:

            '''{extracted_structure}'''
            """

    return generate_slide_json_prompt

def append_response_to_file(api_response, output_file):
    with open(output_file, 'a') as file:
        file.write(api_response)
        file.write("\n")

def iterative_structure_extraction(markdown_content):
    all_slides = []

    sections = acquire_sections(markdown_content)
    
    for i in range(1, len(sections), 2):
        # Step 1: Extract structure from the markdown content
        print(f"Processing section {i}...")
        section_title = sections[i].strip()
        section_content = sections[i+1].strip()
        prompt = step_1_extract_structure(title=section_title, md_content=section_content)
        response = call_openai_api(prompt)
        # print("current response: ", response)
        append_response_to_file(response, 'extracted_structure.txt')

        # Step 2: Generate slide JSON prompt
        # print("section: ", sections[i])
        print("Starting to create slide JSON prompt...")
        slide_json_prompt = step_n_generate_slide_json_prompt(response)
        json_response = call_openai_api(slide_json_prompt)
        print("current json response: ", json_response)
        
        # Step 3: Parse the JSON response and append to all_slides
        try:
            # Extract JSON array from the response
            json_start = json_response.find('[')
            json_end = json_response.rfind(']')
            if json_start != -1 and json_end != -1:
                json_str = json_response[json_start:json_end+1]
                slides = json.loads(json_str)
                all_slides.extend(slides)
                print("all slides: ", all_slides)
            else:
                print(f"JSON not found in the response for section {i//2 + 1}")
                continue  # Skip this section if JSON is not found
        except json.JSONDecodeError as e:
            print(f"JSON decode error in section {i//2 + 1}: {e}")
            continue  # Skip this section if JSON decoding fails

        print("=====================================")

    # Step 4: Write all slides to a JSON file
    with open('slides.json', 'w') as file:
        json.dump(all_slides, file, indent=4)


def iterative_ppt_generation(structured_file):
    pass

if __name__ == "__main__":
    # Load the markdown file
    with open('original_report.md', 'r') as file:
        markdown_content = file.read()

    # Perform iterative structure extraction
    iterative_structure_extraction(markdown_content)
    

    