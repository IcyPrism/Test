from docx import Document
import os
import json

# Set directory and file name
directory = "C:/Users/EUG/Desktop/TEST project"
file_name = "TESTS without answer.docx"
file_path = os.path.join(directory, file_name)

def parse_questions_with_highlights(file_path):
    doc = Document(file_path)
    questions = []
    current_question = {"question": "", "choices": [], "correct": ""}
    in_choices = False

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if text:
            # Detect a new question (starting with a number)
            if text[0].isdigit() and "." in text:
                if current_question["question"]:  # Save the previous question
                    questions.append(current_question)
                    current_question = {"question": "", "choices": [], "correct": ""}
                current_question["question"] = text
                in_choices = True  # Start capturing choices

            # Detect choices (case-sensitive for А., Б., В., Г., Д.)
            elif in_choices and text.startswith(("А.", "Б.", "В.", "Г.", "Д.")):
                # Check if this choice is highlighted
                if "highlight" in paragraph._element.xml:
                    current_question["correct"] = text
                current_question["choices"].append(text)

    # Add the last question if any
    if current_question["question"]:
        questions.append(current_question)

    return questions

# Parse the questions from the file
parsed_questions = parse_questions_with_highlights(file_path)

# Save the results to a JSON file
output_file = os.path.join(directory, "parsed_questions_with_answers1.json")
with open(output_file, "w", encoding="utf-8") as f:
    json.dump(parsed_questions, f, ensure_ascii=False, indent=4)

# Print confirmation
print(f"Questions parsed successfully! Saved to: {output_file}")
