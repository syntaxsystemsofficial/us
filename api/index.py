from flask import Flask, request, send_file, jsonify
import requests
from bs4 import BeautifulSoup
import os
import tempfile

app = Flask(__name__)

@app.route("/api/vin", methods=["GET"])
def get_excel():
    vin = request.args.get("vin")
    if not vin:
        return jsonify({"error": "VIN is required"}), 400

    try:
        url = f"https://vpic.nhtsa.dot.gov/decoder/Decoder?VIN={vin}&ModelYear="
        res = requests.get(url)
        if res.status_code != 200:
            return jsonify({"error": "Failed to load VIN page"}), 500

        soup = BeautifulSoup(res.text, "html.parser")

        # Find the ExportToExcel link
        link_tag = soup.find("a", href=lambda href: href and "ExportToExcel" in href)
        if not link_tag:
            return jsonify({"error": "Excel export link not found"}), 404

        excel_url = "https://vpic.nhtsa.dot.gov" + link_tag["href"]
        excel_response = requests.get(excel_url)

        if excel_response.status_code == 200:
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            temp_file.write(excel_response.content)
            temp_file.close()

            return send_file(
                temp_file.name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=f"{vin}.xlsx"
            )
        else:
            return jsonify({"error": "Failed to download Excel"}), 500

    except Exception as e:
        return jsonify({"error": str(e)}), 500
