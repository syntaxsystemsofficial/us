from flask import Flask, request, send_file, jsonify
import requests
import tempfile

app = Flask(__name__)

@app.route("/api/vin", methods=["GET"])
def get_excel():
    vin = request.args.get("vin")
    if not vin:
        return jsonify({"error": "VIN is required"}), 400

    excel_url = f"https://vpic.nhtsa.dot.gov/decoder/Decoder/ExportToExcel?VIN={vin}"

    try:
        response = requests.get(excel_url)

        if response.status_code == 200 and response.content:
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            temp_file.write(response.content)
            temp_file.close()

            return send_file(
                temp_file.name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=f"{vin}.xlsx"
            )
        else:
            return jsonify({"error": "Excel file not available or VIN is invalid"}), 404

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# For local test
if __name__ == "__main__":
    app.run(debug=True)
