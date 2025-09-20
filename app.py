from flask import Flask, render_template, request, jsonify, send_file
import requests
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import openpyxl
import io

app = Flask(__name__)
GOOGLE_MAPS_API_KEY = "AIzaSyDnL3SwYRLlFwQNpYNdvPYAuc-wU8u_GaA"


def get_geocoded_addresses(addresses):
    """Geocode addresses into formatted addresses."""
    geocoded = []
    for address in addresses:
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {"address": address, "key": GOOGLE_MAPS_API_KEY}
        response = requests.get(url, params=params).json()

        if response['status'] == "OK":
            geocoded.append(response['results'][0]['formatted_address'])
        else:
            raise Exception(f"Could not geocode {address}")
    return geocoded


def get_distance_matrix(locations):
    """Fetch distance & duration matrix from Google."""
    url = "https://maps.googleapis.com/maps/api/distancematrix/json"
    params = {
        "origins": "|".join(locations),
        "destinations": "|".join(locations),
        "key": GOOGLE_MAPS_API_KEY
    }
    response = requests.get(url, params=params).json()

    if response['status'] != "OK":
        raise Exception("Distance matrix request failed")

    n = len(locations)
    dist_matrix = [[0] * n for _ in range(n)]
    time_matrix = [[0] * n for _ in range(n)]

    for i in range(n):
        for j in range(n):
            if i != j:
                dist_matrix[i][j] = response["rows"][i]["elements"][j]["distance"]["value"]  # meters
                time_matrix[i][j] = response["rows"][i]["elements"][j]["duration"]["value"]  # seconds
    return dist_matrix, time_matrix


def solve_tsp(matrix):
    """Simple Nearest Neighbor heuristic for TSP."""
    n = len(matrix)
    visited = [False] * n
    path = [0]  # start at first location
    visited[0] = True

    for _ in range(n - 1):
        last = path[-1]
        nearest = None
        nearest_dist = float("inf")
        for j in range(n):
            if not visited[j] and matrix[last][j] < nearest_dist:
                nearest = j
                nearest_dist = matrix[last][j]
        path.append(nearest)
        visited[nearest] = True

    return path


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/optimize', methods=['POST'])
def optimize():
    start_address = request.form.get('start_address')   # starting point
    delivery_addresses = request.form.getlist('addresses[]')

    try:
        # Step 1: Geocode start + deliveries
        all_addresses = [start_address] + delivery_addresses
        locations = get_geocoded_addresses(all_addresses)

        # Step 2: Distance & Time matrices
        dist_matrix, time_matrix = get_distance_matrix(locations)

        # Step 3: Solve TSP (for deliveries only, index 1..n-1)
        delivery_matrix = [row[1:] for row in dist_matrix[1:]]
        order = solve_tsp(delivery_matrix)  # order of delivery indices (0-based for deliveries)

        # Step 4: Reconstruct full order including start
        optimized_order = [locations[0]] + [locations[i+1] for i in order]

        # Step 5: Calculate total distance & time
        total_distance = 0
        total_time = 0
        full_path = [0] + [i+1 for i in order]  # indices in original matrix
        for i in range(len(full_path) - 1):
            total_distance += dist_matrix[full_path[i]][full_path[i+1]]
            total_time += time_matrix[full_path[i]][full_path[i+1]]

        total_distance_km = round(total_distance / 1000, 2)
        total_time_min = round(total_time / 60, 1)

        return jsonify({
            "optimized_order": optimized_order,
            "total_distance_km": total_distance_km,
            "total_time_min": total_time_min
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/download/pdf', methods=['POST'])
def download_pdf():
    data = request.json
    optimized_order = data.get("optimized_order", [])
    total_distance = data.get("total_distance_km", 0)
    total_time = data.get("total_time_min", 0)

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setFont("Helvetica", 12)

    c.drawString(100, 750, "Delivery Route Plan")
    y = 720
    for i, stop in enumerate(optimized_order, start=1):
        c.drawString(100, y, f"Stop {i}: {stop}")
        y -= 20

    c.drawString(100, y - 10, f"Total Distance: {total_distance} km")
    c.drawString(100, y - 30, f"Estimated Time: {total_time} minutes")

    c.save()
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name="route_plan.pdf", mimetype="application/pdf")


@app.route('/download/excel', methods=['POST'])
def download_excel():
    data = request.json
    optimized_order = data.get("optimized_order", [])
    total_distance = data.get("total_distance_km", 0)
    total_time = data.get("total_time_min", 0)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Route Plan"

    ws.append(["Stop Number", "Address"])
    for i, stop in enumerate(optimized_order, start=1):
        ws.append([i, stop])

    ws.append([])
    ws.append(["Total Distance (km)", total_distance])
    ws.append(["Estimated Time (min)", total_time])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name="route_plan.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    app.run(debug=True)
