from flask import Flask, render_template, send_from_directory, request, abort
import subprocess
import os

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate_ppt/<string:use_case>", methods=["POST"])
def generate_ppt(use_case):
    start_quarter = request.args.get("start_quarter", default=None)
    end_quarter = request.args.get("end_quarter", default=None)
    strategic_theme = request.args.get("strategic_theme", default=None)

    if use_case == "UseCase1":
        script_name = "UseCase1.py"
        output_file = "UseCase1.pptx"
        subprocess.call(["python", script_name])
    elif use_case == "UseCase2":
        script_name = "UseCase2.py"
        output_file = "UseCase2.pptx"
        if start_quarter and end_quarter and strategic_theme:
            subprocess.call(["python", script_name, start_quarter, end_quarter, strategic_theme])
        else:
            abort(400)  # Bad Request
    else:
        abort(404)

    return send_from_directory(
        directory=os.getcwd(),
        path=output_file,
        as_attachment=True,
    )


port = int(os.getenv("PORT", 0))
if __name__ == "__main__":
    if port != 0:
        app.run(host="0.0.0.0", port=port)
    else:
        app.run(debug=True)
