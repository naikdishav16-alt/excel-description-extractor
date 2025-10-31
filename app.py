"""
Flask web app for Excel Description Extractor.

This app allows users to:
1. Upload multiple CSV/Excel files.
2. Process them to extract structured data from the 'Description' column.
3. Download processed results as a ZIP file.
4. Reset (delete) all uploaded and processed files.
"""

from flask import Flask, render_template, request, jsonify, send_file
from processor import process_all_files
import os
import shutil
import zipfile

app = Flask(__name__)

# Define folder paths
INPUT_FOLDER = "Input"
PROCESSED_FOLDER = "processed"
OUTPUT_FOLDER = "output"

# Ensure folders exist
os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/")
def index():
    """
    Renders the main page.

    Returns:
        HTML template: index.html
    """
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload_and_process():
    """
    Handles file upload and triggers processing.

    Returns:
        JSON response: status message for success or failure.
    """
    if "files" not in request.files:
        return jsonify({"status": "error", "message": "No file part"}), 400

    files = request.files.getlist("files")

    if not files or files[0].filename == "":
        return jsonify({"status": "error", "message": "No selected files"}), 400

    for file in files:
        file.save(os.path.join(INPUT_FOLDER, file.filename))

    try:
        process_all_files(INPUT_FOLDER, PROCESSED_FOLDER)

        zip_path = os.path.join(OUTPUT_FOLDER, "processed_files.zip")

        # Create ZIP of processed files
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for root, _, files in os.walk(PROCESSED_FOLDER):
                for file in files:
                    zipf.write(
                        os.path.join(root, file),
                        os.path.relpath(os.path.join(root, file), PROCESSED_FOLDER)
                    )

        return jsonify({"status": "success", "message": "Files processed successfully!"})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/download", methods=["GET"])
def download_zip():
    """
    Allows user to download the ZIP of processed files.

    Returns:
        Flask send_file: ZIP file for download.
    """
    zip_path = os.path.join(OUTPUT_FOLDER, "processed_files.zip")

    if not os.path.exists(zip_path):
        return jsonify({"status": "error", "message": "No processed files found"}), 404

    return send_file(zip_path, as_attachment=True)


@app.route("/delete", methods=["POST"])
def delete_all():
    """
    Deletes all files in Input, Processed, and Output folders.

    Returns:
        JSON response: status message for completion or error.
    """
    try:
        for folder in [INPUT_FOLDER, PROCESSED_FOLDER, OUTPUT_FOLDER]:
            if os.path.exists(folder):
                shutil.rmtree(folder)
                os.makedirs(folder)
        return jsonify({"status": "success", "message": "All files deleted successfully!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)







