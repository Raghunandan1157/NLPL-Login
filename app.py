from flask import Flask, send_from_directory, send_file
import os

app = Flask(__name__, static_folder=None)

ROOT = os.path.dirname(os.path.abspath(__file__))


@app.route("/")
def index():
    return send_file(os.path.join(ROOT, "index.html"))


@app.route("/<path:path>")
def static_files(path):
    return send_from_directory(ROOT, path)


# TODO: POST /api/upload — accept 2-file Excel upload for processing


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
