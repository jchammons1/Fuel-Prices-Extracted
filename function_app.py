import azure.functions as func
import logging
import os
import tempfile
import base64
import json

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="fuel_prices_extracted")
def fuel_prices_extracted(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Fuel prices extraction function started.")

    try:
        from main import process_fuel_prices

        req_body = req.get_json()

        file_name = req_body.get("fileName")
        file_content = req_body.get("fileContent")

        if not file_name or not file_content:
            return func.HttpResponse(
                "Missing fileName or fileContent.",
                status_code=400
            )

        # Decode Base64 PDF content
        file_bytes = base64.b64decode(file_content)

        # Save to temp file
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, file_name)

        with open(temp_path, "wb") as f:
            f.write(file_bytes)

        # Parse PDF and return extracted values
        result = process_fuel_prices(temp_path)

        return func.HttpResponse(
            json.dumps(result),
            mimetype="application/json",
            status_code=200
        )

    except Exception as e:
        logging.exception("Error running fuel price extraction.")
        return func.HttpResponse(
            f"Error: {str(e)}",
            status_code=500
        )