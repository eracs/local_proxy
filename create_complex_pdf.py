from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import os
import math
import random

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))


def create_gradient_pattern(filename, width=800, height=400):
    img = Image.new("RGB", (width, height))
    draw = ImageDraw.Draw(img)
    for y in range(height):
        r = int(30 + (200 - 30) * (y / height))
        g = int(60 + (120 - 60) * (y / height))
        b = int(150 + (250 - 150) * (y / height))
        draw.line([(0, y), (width, y)], fill=(r, g, b))
    for i in range(0, width, 40):
        alpha = int(255 * (i / width))
        draw.line([(i, 0), (i + 100, height)], fill=(255, 255, 255, alpha // 10), width=2)
    for _ in range(50):
        x = random.randint(0, width)
        y = random.randint(0, height)
        size = random.randint(2, 8)
        draw.ellipse([x, y, x + size, y + size], fill=(255, 255, 255, 180))
    path = os.path.join(OUTPUT_DIR, filename)
    img.save(path)
    return path


def create_mandelbrot_fractal(filename, width=600, height=600):
    img = Image.new("RGB", (width, height))
    pixels = img.load()
    for py in range(height):
        for px in range(width):
            x0 = (px / width) * 3.5 - 2.5
            y0 = (py / height) * 2.0 - 1.0
            x, y = 0.0, 0.0
            iteration = 0
            max_iter = 100
            while x * x + y * y <= 4 and iteration < max_iter:
                xtemp = x * x - y * y + x0
                y = 2 * x * y + y0
                x = xtemp
                iteration += 1
            if iteration == max_iter:
                pixels[px, py] = (0, 0, 0)
            else:
                r = int((iteration / max_iter) * 255)
                g = int((iteration / max_iter) * 150)
                b = int(255 - (iteration / max_iter) * 200)
                pixels[px, py] = (r, g, b)
    path = os.path.join(OUTPUT_DIR, filename)
    img.save(path)
    return path


def create_multi_line_chart(filename):
    width, height = 700, 400
    img = Image.new("RGB", (width, height), "#f8f9fa")
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 14)
        title_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 20)
    except (OSError, IOError):
        font = ImageFont.load_default()
        title_font = font
    draw.text((250, 15), "Multi-Series Performance", fill="#2c3e50", font=title_font)
    margin_left, margin_right = 60, 30
    margin_top, margin_bottom = 60, 50
    chart_width = width - margin_left - margin_right
    chart_height = height - margin_top - margin_bottom
    draw.rectangle([margin_left, margin_top, width - margin_right, height - margin_bottom], outline="#dee2e6", width=2)
    for i in range(6):
        y = margin_top + (chart_height * i) // 5
        draw.line([(margin_left, y), (width - margin_right, y)], fill="#e9ecef", width=1)
        draw.text((margin_left - 40, y - 8), str(100 - i * 20), fill="#6c757d", font=font)
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
    for i, month in enumerate(months):
        x = margin_left + (chart_width * i) // (len(months) - 1)
        draw.text((x - 15, height - margin_bottom + 10), month, fill="#6c757d", font=font)
    datasets = [
        {"label": "Product A", "data": [65, 72, 78, 75, 80, 85], "color": "#e74c3c"},
        {"label": "Product B", "data": [45, 55, 60, 68, 72, 78], "color": "#3498db"},
        {"label": "Product C", "data": [30, 38, 42, 50, 58, 65], "color": "#2ecc71"},
    ]
    for dataset in datasets:
        points = []
        for i, value in enumerate(dataset["data"]):
            x = margin_left + (chart_width * i) // (len(months) - 1)
            y = margin_top + chart_height - (value / 100) * chart_height
            points.append((x, y))
        for i in range(len(points) - 1):
            draw.line([points[i], points[i + 1]], fill=dataset["color"], width=3)
        for x, y in points:
            draw.ellipse([x - 5, y - 5, x + 5, y + 5], fill=dataset["color"], outline="white", width=2)
    legend_y = height - 25
    for i, dataset in enumerate(datasets):
        legend_x = margin_left + i * 150
        draw.rectangle([legend_x, legend_y, legend_x + 20, legend_y + 12], fill=dataset["color"])
        draw.text((legend_x + 25, legend_y - 2), dataset["label"], fill="#2c3e50", font=font)
    path = os.path.join(OUTPUT_DIR, filename)
    img.save(path)
    return path


def create_spiral_pattern(filename, width=600, height=600):
    img = Image.new("RGB", (width, height), "#1a1a2e")
    draw = ImageDraw.Draw(img)
    center_x, center_y = width // 2, height // 2
    for angle in range(0, 3600, 5):
        rad = math.radians(angle)
        radius = angle * 0.08
        x = center_x + radius * math.cos(rad)
        y = center_y + radius * math.sin(rad)
        if 0 <= x < width and 0 <= y < height:
            r = int(128 + 127 * math.sin(rad))
            g = int(128 + 127 * math.sin(rad + 2 * math.pi / 3))
            b = int(128 + 127 * math.sin(rad + 4 * math.pi / 3))
            size = max(2, int(8 - angle / 500))
            draw.ellipse([x - size, y - size, x + size, y + size], fill=(r, g, b))
    path = os.path.join(OUTPUT_DIR, filename)
    img.save(path)
    return path


def create_pie_chart(filename):
    width, height = 500, 500
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 16)
        title_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 22)
    except (OSError, IOError):
        font = ImageFont.load_default()
        title_font = font
    draw.text((180, 20), "Market Share Distribution", fill="#2c3e50", font=title_font)
    data = [("Product A", 35, "#e74c3c"), ("Product B", 25, "#3498db"), ("Product C", 20, "#2ecc71"), ("Product D", 12, "#f39c12"), ("Others", 8, "#9b59b6")]
    total = sum(d[1] for d in data)
    cx, cy, radius = 250, 270, 150
    start_angle = 0
    for label, value, color in data:
        end_angle = start_angle + (value / total) * 360
        draw.pieslice([cx - radius, cy - radius, cx + radius, cy + radius], start_angle, end_angle, fill=color, outline="white", width=3)
        mid_angle = math.radians((start_angle + end_angle) / 2)
        label_x = cx + (radius + 30) * math.cos(mid_angle)
        label_y = cy + (radius + 30) * math.sin(mid_angle)
        draw.text((label_x - 20, label_y - 10), f"{label}\n{value}%", fill="#2c3e50", font=font)
        start_angle = end_angle
    path = os.path.join(OUTPUT_DIR, filename)
    img.save(path)
    return path


def create_geometric_pattern(filename, width=700, height=500):
    img = Image.new("RGB", (width, height), "#0f0f23")
    draw = ImageDraw.Draw(img)
    for i in range(0, width, 50):
        for j in range(0, height, 50):
            hue = (i + j) % 360
            r = int(128 + 127 * math.sin(math.radians(hue)))
            g = int(128 + 127 * math.sin(math.radians(hue + 120)))
            b = int(128 + 127 * math.sin(math.radians(hue + 240)))
            shape_type = (i // 50 + j // 50) % 3
            if shape_type == 0:
                draw.rectangle([i + 5, j + 5, i + 45, j + 45], fill=(r, g, b), outline="white", width=2)
            elif shape_type == 1:
                draw.ellipse([i + 5, j + 5, i + 45, j + 45], fill=(r, g, b), outline="white", width=2)
            else:
                draw.polygon([(i + 25, j + 5), (i + 45, j + 45), (i + 5, j + 45)], fill=(r, g, b), outline="white", width=2)
    path = os.path.join(OUTPUT_DIR, filename)
    img.save(path)
    return path


print("Generating complex graphics...")
img_gradient = create_gradient_pattern("gradient.png")
img_fractal = create_mandelbrot_fractal("fractal.png")
img_chart = create_multi_line_chart("multiline_chart.png")
img_spiral = create_spiral_pattern("spiral.png")
img_pie = create_pie_chart("pie_chart.png")
img_geometric = create_geometric_pattern("geometric.png")
print("All images generated.")


class PDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 18)
        self.cell(0, 12, "Advanced Graphics in PDF", align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(3)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 9)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")


pdf = PDF()
pdf.alias_nb_pages()
pdf.set_auto_page_break(auto=True, margin=20)

# Page 1: Gradient banner + geometric pattern
pdf.add_page()
pdf.set_font("Helvetica", "", 11)
pdf.multi_cell(0, 6, "This PDF showcases complex graphics generated with Python's Pillow library, including gradients, fractals, charts, and geometric patterns.")
pdf.ln(5)
pdf.image(img_gradient, x=10, w=190)
pdf.ln(4)
pdf.set_font("Helvetica", "I", 9)
pdf.cell(0, 5, "Figure 1: Gradient background with overlay patterns", align="C")
pdf.ln(8)
pdf.image(img_geometric, x=15, w=180)
pdf.ln(4)
pdf.set_font("Helvetica", "I", 9)
pdf.cell(0, 5, "Figure 2: Geometric tessellation with dynamic colors", align="C")

# Page 2: Mandelbrot fractal
pdf.add_page()
pdf.set_font("Helvetica", "B", 14)
pdf.cell(0, 10, "Fractal Mathematics", new_x="LMARGIN", new_y="NEXT")
pdf.ln(3)
pdf.set_font("Helvetica", "", 11)
pdf.multi_cell(0, 6, "The Mandelbrot set is a famous fractal generated by iterating the formula z = z² + c. Each pixel's color represents how quickly the iteration diverges.")
pdf.ln(5)
pdf.image(img_fractal, x=35, w=140)
pdf.ln(4)
pdf.set_font("Helvetica", "I", 9)
pdf.cell(0, 5, "Figure 3: Mandelbrot set visualization (100 iterations)", align="C")

# Page 3: Multi-line chart
pdf.add_page()
pdf.set_font("Helvetica", "B", 14)
pdf.cell(0, 10, "Data Analytics Dashboard", new_x="LMARGIN", new_y="NEXT")
pdf.ln(3)
pdf.image(img_chart, x=10, w=190)
pdf.ln(4)
pdf.set_font("Helvetica", "I", 9)
pdf.cell(0, 5, "Figure 4: Multi-series line chart with interactive-style design", align="C")
pdf.ln(8)
pdf.set_font("Helvetica", "", 11)
pdf.multi_cell(0, 6, "The chart above demonstrates multiple data series with distinct colors, data point markers, grid lines, and a legend - all generated programmatically.")

# Page 4: Pie chart + spiral
pdf.add_page()
pdf.set_font("Helvetica", "B", 14)
pdf.cell(0, 10, "Visual Patterns", new_x="LMARGIN", new_y="NEXT")
pdf.ln(3)
pdf.image(img_pie, x=10, y=35, w=90)
pdf.image(img_spiral, x=105, y=35, w=90)
pdf.set_xy(10, 150)
pdf.set_font("Helvetica", "I", 9)
pdf.cell(90, 5, "Figure 5: Pie chart with labels", align="C")
pdf.cell(95, 5, "Figure 6: Spiral pattern with RGB color cycling", align="C")
pdf.ln(10)
pdf.set_font("Helvetica", "", 11)
pdf.multi_cell(0, 6, "Left: Market share distribution shown as a pie chart with percentage labels. Right: A mathematical spiral pattern where colors cycle through the RGB spectrum based on angle.")

pdf_path = os.path.join(OUTPUT_DIR, "complex_graphics.pdf")
pdf.output(pdf_path)
print(f"\nPDF created: {pdf_path}")

for f in ["gradient.png", "fractal.png", "multiline_chart.png", "spiral.png", "pie_chart.png", "geometric.png"]:
    os.remove(os.path.join(OUTPUT_DIR, f))
print("Temp images cleaned up.")
