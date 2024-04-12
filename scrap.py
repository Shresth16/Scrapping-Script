import os
import openpyxl
import json


def read_data_from_excel(excel_file):
    extracted_data = []
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active

    for row in sheet.iter_rows(values_only=True):
        # Assuming the first column contains text data
        text_data = row[0]

        # Assuming the second column contains image URLs
        images = []  # Placeholder for image URLs
        # Assuming the third column contains link URLs
        links = []  # Placeholder for link URLs

        # Append data to extracted_data
        page_data = {
            'text': text_data,
            'images': images,
            'links': links
        }
        extracted_data.append(page_data)

    return extracted_data


def save_data_to_json(data, output_file):
    with open(output_file, 'w') as json_file:
        json.dump(data, json_file, indent=4)


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_file = os.path.join(script_dir,"C:\\Users\\shres\\Downloads\\Scrapping Python Assigment- Flair Insights.xlsx")
    output_file = "scraped_data.json"

    extracted_data = read_data_from_excel(excel_file)
    save_data_to_json(extracted_data, output_file)
