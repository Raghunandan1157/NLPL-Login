from flask import Flask, send_from_directory, send_file, request, jsonify
import os
from raw_parser import parse_raw_file

app = Flask(__name__, static_folder=None)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max upload

ROOT = os.path.dirname(os.path.abspath(__file__))


@app.route("/")
def index():
    return send_file(os.path.join(ROOT, "index.html"))


@app.route("/api/upload-raw", methods=['POST'])
def upload_raw():
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'No file provided'}), 400

    f = request.files['file']
    if not f.filename:
        return jsonify({'status': 'error', 'message': 'No file selected'}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.xlsx', '.xlsm'):
        return jsonify({'status': 'error', 'message': 'Only .xlsx files are accepted'}), 400

    try:
        file_bytes = f.read()
        try:
            result = parse_raw_file(file_bytes, include_accounts=True)
        except MemoryError:
            # Fallback: skip account-level detail to save memory
            result = parse_raw_file(file_bytes, include_accounts=False)
            result['meta']['fallback'] = 'officer_only'

        del file_bytes  # free memory
        return jsonify(result)
    except ValueError as e:
        return jsonify({'status': 'error', 'message': str(e)}), 400
    except Exception as e:
        return jsonify({'status': 'error', 'message': 'Failed to process file: ' + str(e)}), 500


@app.route("/<path:path>")
def static_files(path):
    return send_from_directory(ROOT, path)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
