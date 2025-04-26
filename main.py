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
    dropped_files = sys.argv[1:]

    if not dropped_files:
        print("--------------------------------------------------")
        print("Usage:")
        print("  Drag and drop one PDF file or one or more PNG files")
        print("  onto this executable.")
        print("--------------------------------------------------")
        input("Press any key to exit...")
        sys.exit(0)

    first_file_path = dropped_files[0]
    file_name, file_ext = os.path.splitext(first_file_path)
    file_ext = file_ext.lower()

    success = False
    if file_ext == ".pdf":
        # --- PDF -> PNG Mode ---
        if len(dropped_files) > 1:
            print("Error: For PDF to PNG conversion, please drag and drop only one PDF file at a time.")
        elif not os.path.exists(first_file_path):
            print(f"Error: The specified PDF file was not found: {first_file_path}")
        else:
            output_dir = os.path.join(os.path.dirname(first_file_path), os.path.basename(file_name) + "_png")
            success = convert_pdf_to_png(first_file_path, output_dir)

    elif file_ext == ".png":
        # --- PNG -> PDF Mode ---
        image_paths_unsorted = []
        valid_pngs = True
        for f_path in dropped_files:
            if not os.path.splitext(f_path)[1].lower() == ".png":
                print(f"Error: '{os.path.basename(f_path)}' is not a PNG file. Please drag and drop only PNG files.")
                valid_pngs = False
                break
            elif not os.path.exists(f_path):
                print(f"Error: The specified PNG file was not found: {f_path}")
                valid_pngs = False
                break
            else:
                image_paths_unsorted.append(f_path)

        if valid_pngs and image_paths_unsorted:
            # <<< Sort the file path list using natural sort >>>
            print("Sorting files by name (natural sort)...")
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
            # Remove trailing numbers and common separators (like _ or -)
            # This will attempt to find a common prefix for the filenames
            output_prefix = re.sub(r'[_-]?\d+$', '', first_image_base)
            # If the prefix becomes empty (e.g., filename was just "1.png"), use a default
            if not output_prefix:
                output_prefix = "combined"

            output_dir_path = os.path.dirname(image_paths[0])
            output_base_name = os.path.join(output_dir_path, output_prefix)
            output_pdf_path = output_base_name + ".pdf"
            count = 1
            # Handle potential filename conflicts
            while os.path.exists(output_pdf_path):
                output_pdf_path = f"{output_base_name}_{count}.pdf"
                count += 1

            # Pass the sorted image_paths for PDF conversion
            success = convert_png_to_pdf(image_paths, output_pdf_path)

    else:
        print(f"Error: Unsupported file format: '{os.path.basename(first_file_path)}'")
        print("Please drag and drop a PDF file (.pdf) or PNG files (.png).")

    # Display message based on processing result and wait before closing the window
    if success:
        print("Processing completed successfully.")
    else:
        print("An error occurred during processing.")

    input("Press any key to exit...")