import fitz  # PyMuPDF
import os
import sys
import re  # Import for natural sort
from PIL import Image
import io

# --- Helper function for natural sort ---
def natural_sort_key(s):
    """Generate a key for natural sort (e.g., 'page1', 'page2', 'page10')."""
    # Split the string into tuples of non-digit and digit parts
    # Example: "image10.png" -> ('image', 10, '.png')
    # This ensures that the numeric parts are compared correctly
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', os.path.basename(s))] # Sort by filename part only (Changed regex to split by digit)

def convert_pdf_to_png(pdf_path, output_dir):
    """Convert each page of a PDF file into a PNG image."""
    try:
        doc = fitz.open(pdf_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        print(f"Converting '{os.path.basename(pdf_path)}' to PNG...")
        pdf_base_name = os.path.splitext(os.path.basename(pdf_path))[0] # Get PDF base name
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            # 解像度を上げるためにマトリックスの値を変更 (例: 4x4 = 288 DPI)
            # 値を大きくすると画質は向上するが、ファイルサイズも大きくなる
            mat = fitz.Matrix(4, 4)
            pix = page.get_pixmap(matrix=mat)
            # Use PDF base name for output filename
            output_filename = f"{pdf_base_name}_{page_num + 1}.png"
            output_path = os.path.join(output_dir, output_filename)
            pix.save(output_path)
            print(f"  - Saved: {output_filename}")
        doc.close()
        print(f"Conversion to PNG completed. Output directory: {output_dir}")
        return True
    except Exception as e:
        print(f"Error (PDF -> PNG): {e}")
        return False

def convert_png_to_pdf(image_paths, output_pdf_path):
    """Convert multiple PNG images into a single PDF file (in the specified order)."""
    try:
        print(f"Converting PNG images to PDF (page order determined)...")
        doc = fitz.open() # Create a new empty PDF
        for i, img_path in enumerate(image_paths):
            print(f"  - Page {i+1}: '{os.path.basename(img_path)}'")
            try:
                img = fitz.open(img_path) # Open the image file
                rect = img[0].rect # Get the size of the image
                # Create a page with the image size
                pdf_page = doc.new_page(width=rect.width, height=rect.height)
                pdf_page.insert_image(rect, filename=img_path) # Insert the image onto the page
                img.close()
            except Exception as page_e:
                print(f"    Error: An error occurred while processing '{os.path.basename(img_path)}': {page_e}")
                print("    This file will be skipped.")
                # If processing should continue even if an error occurs on a page.
                # To stop processing, use raise page_e etc.

        if len(doc) > 0: # Save only if at least one page was processed
            doc.save(output_pdf_path)
            doc.close()
            print(f"Conversion to PDF completed. Output file: {output_pdf_path}")
            return True
        else:
            print("Error: No valid PNG images were found, so the PDF was not created.")
            doc.close()
            return False

    except Exception as e:
        print(f"Error (PNG -> PDF): {e}")
        # Close doc as it might be open
        try:
            if 'doc' in locals() and doc:
                doc.close()
        except:
            pass # Ignore errors during error handling
        return False


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("--------------------------------------------------")
        print("Usage:")
        print(f"  python {os.path.basename(sys.argv[0])} <target_folder>")
        print("  or drag and drop a FOLDER onto the executable.")
        print("\nDescription:")
        print("  Processes files within the specified <target_folder>.")
        print("  - If exactly one PDF file is found, it's converted to PNGs.")
        print("  - If multiple PNG files are found, they are combined into a single PDF.")
        print("--------------------------------------------------")
        input("Press any key to exit...")
        sys.exit(1)

    target_folder = sys.argv[1]

    if not os.path.isdir(target_folder):
        print(f"Error: The specified path is not a valid directory: {target_folder}")
        input("Press any key to exit...")
        sys.exit(1)

    # Define and create the base output directory
    output_base_dir = "output"
    try:
        os.makedirs(output_base_dir, exist_ok=True)
        print(f"Output will be saved to: {os.path.abspath(output_base_dir)}")
    except Exception as e:
        print(f"Error creating output directory '{output_base_dir}': {e}")
        input("Press any key to exit...")
        sys.exit(1)

    pdf_files = []
    png_files = []
    print(f"Scanning folder: {target_folder}")
    try:
        for filename in os.listdir(target_folder):
            file_path = os.path.join(target_folder, filename)
            if os.path.isfile(file_path):
                _, file_ext = os.path.splitext(filename)
                file_ext = file_ext.lower()
                if file_ext == ".pdf":
                    pdf_files.append(file_path)
                elif file_ext == ".png":
                    png_files.append(file_path)
    except Exception as e:
        print(f"Error reading directory contents: {e}")
        input("Press any key to exit...")
        sys.exit(1)

    overall_success = True # Track success across multiple operations
    processed_something = False # Flag to check if any operation was performed
    num_pdf = len(pdf_files)
    num_png = len(png_files)

    # --- Determine operation mode based on folder content ---
    if num_pdf >= 1 and num_png == 0:
        # --- PDF -> PNG Mode (Process all found PDFs) ---
        print(f"Mode: PDF -> PNG (Found {num_pdf} PDF files)")
        processed_something = True
        for pdf_path in pdf_files:
            pdf_base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            # Output directory inside ./output
            output_dir = os.path.join(output_base_dir, pdf_base_name + "_png")
            print(f"\nProcessing '{os.path.basename(pdf_path)}'...")
            success = convert_pdf_to_png(pdf_path, output_dir)
            if not success:
                overall_success = False # Mark failure if any conversion fails

    elif num_png > 1 and num_pdf == 0:
        # --- PNG -> PDF Mode ---
        print(f"Mode: PNG -> PDF (Found {num_png} PNG files)")
        processed_something = True
        image_paths_unsorted = png_files

        # <<< Sort the file path list using natural sort >>>
        print("Sorting PNG files by name (natural sort)...")
        try:
            image_paths = sorted(image_paths_unsorted, key=natural_sort_key)
            print("Sorted file order (this will be the PDF page order):")
            for i, p in enumerate(image_paths):
                print(f"  {i+1}: {os.path.basename(p)}")
        except Exception as sort_e:
            print(f"Error: An error occurred during file sorting: {sort_e}")
            print("Continuing without sorting (will use the order provided by the OS).")
            image_paths = image_paths_unsorted # Use original order on error

        # Determine the output PDF filename based on the first image's name pattern
        first_image_name = os.path.basename(image_paths[0])
        first_image_base, _ = os.path.splitext(first_image_name)
        output_prefix = re.sub(r'[_-]?\\d+$', '', first_image_base)
        if not output_prefix:
            output_prefix = "combined" # Default name if prefix is empty

        # Output PDF inside ./output
        output_base_name = os.path.join(output_base_dir, output_prefix)
        output_pdf_path = output_base_name + ".pdf"
        count = 1
        # Handle potential filename conflicts
        while os.path.exists(output_pdf_path):
            output_pdf_path = f"{output_base_name}_{count}.pdf"
            count += 1

        # Pass the sorted image_paths for PDF conversion
        success = convert_png_to_pdf(image_paths, output_pdf_path)
        if not success:
            overall_success = False

    else:
        # --- Invalid file combination or no files ---
        processed_something = False # No valid operation determined
        print("Error: Invalid file combination or no processable files found in the target folder.")
        if num_pdf > 0 and num_png > 0:
            print("  - Found both PDF and PNG files. Please provide a folder with only PDFs or only PNGs.")
            print(f"    PDF(s): {[os.path.basename(f) for f in pdf_files]}")
            print(f"    PNG(s): {[os.path.basename(f) for f in png_files]}")
        elif num_pdf == 0 and num_png == 1:
             print("  - Found only one PNG file. Need multiple PNGs to combine into a PDF.")
             print(f"    PNG: {os.path.basename(png_files[0])}")
        elif num_pdf == 0 and num_png == 0:
            print("  - No PDF or PNG files found in the specified folder.")
        # The case num_pdf > 1 and num_png == 0 is now handled above
        overall_success = False # Mark as failure

    # Display message based on processing result
    print("\n--------------------------------------------------")
    if processed_something:
        if overall_success:
            print("Processing completed successfully.")
        else:
            print("Processing completed, but some errors occurred.")
    else:
        # Error message was already printed in the 'else' block above
        print("No processing was performed due to errors or invalid folder content.")
    print(f"Check the '{os.path.abspath(output_base_dir)}' directory for results.")
    print("--------------------------------------------------")

    input("Press any key to exit...")