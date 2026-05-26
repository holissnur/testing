import os
import sys
import tkinter as tk
from tkinter import ttk
import fitz  # PyMuPDF


def process_pdfs_with_gui(master_file, template_files, target_fields):
    # 1. Create Output Folder if it doesn't exist
    output_folder = "Output"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 2. Setup Tkinter GUI Window
    root = tk.Tk()
    root.title("Copy To Another")
    root.geometry("400x190")  # Expanded height slightly for the button
    root.resizable(False, False)

    # Status Label
    status_label = tk.Label(
        root, text="Initializing...", font=("Arial", 10, "bold"), anchor="w"
    )
    status_label.pack(fill="x", padx=20, pady=(15, 5))

    # Progress Bar Component
    progress_bar = ttk.Progressbar(
        root, orient="horizontal", length=360, mode="determinate"
    )
    progress_bar.pack(padx=20, pady=5)
    progress_bar["maximum"] = len(template_files)

    # Detail Label (shows current file or output path)
    detail_label = tk.Label(
        root, text="", font=("Arial", 9), fg="gray", anchor="w"
    )
    detail_label.pack(fill="x", padx=20, pady=(0, 10))

    # Force UI to draw before processing starts
    root.update()

    try:
        # 3. Read values from the master PDF
        status_label.config(text="Reading Master PDF fields...", fg="black")
        root.update()

        form_values = {}
        src_doc = fitz.open(master_file)
        for page in src_doc:
            for field in page.widgets():
                if field.field_name in target_fields:
                    form_values[field.field_name] = field.field_value
        src_doc.close()

        # 4. Copy values to Template PDFs and update GUI
        for index, target_path in enumerate(template_files, start=1):
            status_label.config(
                text=f"Processing file {index} of {len(template_files)}..."
            )
            detail_label.config(text=f"Current: {os.path.basename(target_path)}")
            root.update()

            tgt_doc = fitz.open(target_path)
            for page in tgt_doc:
                for field in page.widgets():
                    if field.field_name in form_values:
                        field.field_value = form_values[field.field_name]
                        field.update()

            # Save file inside the designated Output folder
            base_name = os.path.basename(target_path)
            output_path = os.path.join(
                output_folder, base_name.replace(".pdf", "_Filled.pdf")
            )

            tgt_doc.save(
                output_path, incremental=False, encryption=fitz.PDF_ENCRYPT_KEEP
            )
            tgt_doc.close()

            # Update progress bar positions
            progress_bar["value"] = index
            root.update()

        # 5. Integrated Status Display
        status_label.config(text="Success: Processing Complete!", fg="green")
        detail_label.config(
            text=f"Saved {len(template_files)} files to folder: ./{output_folder}/",
            fg="black",
        )

    except Exception as e:
        # Display errors inline inside the GUI status bar using red text
        status_label.config(text="Error Occurred!", fg="red")
        detail_label.config(text=str(e), fg="red")

    # 6. DYNAMIC CLOSE BUTTON (Appears only after finish/error)
    close_button = tk.Button(
        root, text="OK", command=root.destroy, width=15
    )
    close_button.pack(pady=(5, 10))

    # Keep the GUI window open so the user can click close
    root.mainloop()


# Run Configurations
master_file = "1_Master_Filled.pdf"
template_files = ["2_John_Template.pdf", "3_Jane_Template.pdf"]
target_fields = ["Address", "State", "Type Of Registration"]

if __name__ == "__main__":
    process_pdfs_with_gui(master_file, template_files, target_fields)
