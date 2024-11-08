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

with open('prompts.yaml', 'r') as file:
    api_prompts = yaml.safe_load(file)

# Function to send a prompt to OpenAI API and get a response
def call_openai_api(prompt):
    response = client.chat.completions.create(model=api_prompts['inference_parameters']['model'],
    messages=[
      {"role": "system", "content": "You are an assistant that converts markdown content into a detailed, structured, and visually appealing powerpoint presentation."},
      {"role": "user", "content": prompt}
    ],
    max_tokens=api_prompts['inference_parameters']['max_token'],
    temperature=api_prompts['inference_parameters']['temperature'],)
    return response.choices[0].message.content

def save_api_response(response, output_file):
    with open(output_file, 'w') as file:
        file.write(response)

def acquire_sections(markdown_file):
    sections = re.split(r'(^#+ .*$)', markdown_file, flags=re.MULTILINE)
    return sections

def step_1_extract_structure(md_content, title="Extract Markdown Structure"):
    # Prepare the prompt
    extract_structure_prompt = api_prompts['extract_structure_prompt'].format(title=title, md_content=md_content)
    return extract_structure_prompt

def step_n_generate_slide_json_prompt(extracted_structure):
    # Prepare the prompt
    generate_slide_json_prompt = api_prompts['generate_slide_json_prompt'].format(extracted_structure=extracted_structure)
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
    with open('latest_report.md', 'r') as file:
        markdown_content = file.read()

    # Perform iterative structure extraction
    iterative_structure_extraction(markdown_content)
    

    