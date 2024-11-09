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
    sections = re.split(r'(^## .*$)', markdown_file, flags=re.MULTILINE)
    return sections

def step_1_extract_structure(md_content, title="Extract Markdown Structure"):
    # Prepare the prompt
    extract_structure_prompt = api_prompts['extract_structure_prompt'].format(title=title, md_content=md_content)
    return extract_structure_prompt

def step_2_generate_slide_json_prompt(extracted_structure):
    # Prepare the prompt
    generate_slide_json_prompt = api_prompts['generate_slide_json_prompt'].format(extracted_structure=extracted_structure)
    return generate_slide_json_prompt

def step_3_condense_slide_json_prompt(slide_json):
    # Prepare the prompt
    condense_slide_json_prompt = api_prompts['generate_combined_slide_prompt'].format(slides_chunk=slide_json)
    return condense_slide_json_prompt

def append_response_to_file(api_response, output_file):
    with open(output_file, 'a') as file:
        file.write(api_response)
        file.write("\n")

def iterative_structure_extraction(markdown_content):
    all_slides = []

    sections = acquire_sections(markdown_content)
    
    with open('sections.txt', 'w') as file:
        for i in range(1, len(sections), 2):
            file.write(sections[i].strip())
            file.write("\n")
    
    for i in range(1, len(sections), 2):
        # Step 1: Extract structure from the markdown content
        print(f"Processing section {i // 2 + 1}...")
        section_title = sections[i].strip()
        section_content = sections[i+1].strip()
        prompt = step_1_extract_structure(title=section_title, md_content=section_content)
        response = call_openai_api(prompt)
        # print("current response: ", response)
        append_response_to_file(response, 'extracted_structure.txt')

        # Step 2: Generate slide JSON prompt
        # print("section: ", sections[i])
        print("Starting to create slide JSON prompt...")
        slide_json_prompt = step_2_generate_slide_json_prompt(response)
        json_response = call_openai_api(slide_json_prompt)
        # print("current json response: ", json_response)
        
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

def split_json_into_chunks(original_json, chunk_size):
    return [original_json[i:i + chunk_size] for i in range(0, len(original_json), chunk_size)]

def side_condensation(original_json, target_slide_count):
    json_chunk = split_json_into_chunks(original_json, 10)
    condensed_slides = []

    for chunk in json_chunk:
        print("Condensing chunk...")
        prompt = step_3_condense_slide_json_prompt(chunk)
        response = call_openai_api(prompt)
        print("current response: ", response)
        print("current response in JSON: ", json.loads(response))
        condensed_slides.extend(json.loads(response))
        print("current condensed slides: ", condensed_slides)

    # Further condense if we still exceed target slide count
    if len(condensed_slides) > target_slide_count:
        final_chunk = split_json_into_chunks(condensed_slides, len(condensed_slides) // target_slide_count)
        final_condensed = []
        for chunk in final_chunk:
            print("Further condensing chunk: ", chunk)
            prompt = step_3_condense_slide_json_prompt(chunk)
            response = call_openai_api(prompt)
            final_condensed.extend(json.loads(response))
        condensed_slides = final_condensed[:target_slide_count]

    with open('condensed_slides.json', 'w') as file:
        json.dump(condensed_slides[:target_slide_count], file, indent=4)

    # return condensed_slides[:target_slide_count]

if __name__ == "__main__":
    # Load the markdown file
    with open('latest_report.md', 'r') as file:
        markdown_content = file.read()

    with open('slides.json', 'r') as file:
        original_json = json.load(file)

    # with open('temp_chunks.json', 'r') as file:
    #     chunk_json = json.load(file)

    # for each in chunk_json:
    #     print(each)
    #     print("=====================================")

    # Perform iterative structure extraction
    # iterative_structure_extraction(markdown_content)

    # Perform side condensation
    side_condensation(original_json, 10)
    

    