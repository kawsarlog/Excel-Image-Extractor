import os
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import configparser
from PIL import Image

class ExcelImageProcessor:
    def __init__(self, config_file):
        # Initialize the class with configuration settings
        self.config = configparser.ConfigParser()
        self.config.read(config_file)
        
        # Define output folder and report filename from config
        self.output_folder = self.config['Output']['folder_location']
        self.report_filename = self.config['Output']['report_filename']

        # Create the output folder if it doesn't exist
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

        # Load the Excel workbook and sheet
        self.pxl_doc = openpyxl.load_workbook(self.config['Excel']['filename'])
        self.sheet = self.pxl_doc[self.config['Excel']['sheetname']]
        self.image_loader = SheetImageLoader(self.sheet)
        self.last_row = self.sheet.max_row

    def process_images(self):
        successful_data = []  # List to store successful processing data
        unsuccessful_data = []  # List to store unsuccessful processing data
        errors = []  # List to store encountered errors

        # Iterate through rows, get image and text, and save the images
        for row_number in range(2, self.last_row + 1):
            image_cell = f"{self.config['Columns']['image_column']}{row_number}"
            text_cell = f"{self.config['Columns']['text_column']}{row_number}"

            try:
                image = self.image_loader.get(image_cell)

                if image:
                    image_rgb = image.convert('RGB')
                    cell_text = self.sheet[text_cell].value
                    image_filename = f"{cell_text}.jpg"
                    image_path = os.path.join(self.output_folder, image_filename)
                    image_rgb.save(image_path)
                    successful_data.append(f"Image Cell: {image_cell} -> Saved as: {image_filename}")
                else:
                    unsuccessful_data.append(f"Not saved: {image_cell} (Text Cell: {text_cell})")

            except Exception as e:
                errors.append(f"Error processing row {row_number}: {str(e)}")

        # Write the report to a text file
        report_path = os.path.join(self.output_folder, self.report_filename)
        with open(report_path, 'w') as report_file:
            report_file.write("Successful Data:\n")
            report_file.write("\n".join(successful_data))
            report_file.write("\n\nUnsuccessful Data:\n")
            report_file.write("\n".join(unsuccessful_data))
            report_file.write("\n\nErrors:\n")
            report_file.write("\n".join(errors))

        print("Processing completed. Report generated at:", report_path)

# Main section
if __name__ == "__main__":
    # Configuration file path
    config_file = 'config.ini'
    
    # Initialize the processor and trigger image processing
    processor = ExcelImageProcessor(config_file)
    processor.process_images()
