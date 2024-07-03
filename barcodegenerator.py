import barcode
from barcode import Code128
from barcode.writer import ImageWriter

def generate_barcode(id_number, format="png"):

    my_code = Code128(id_number, writer=ImageWriter())
    filename = f"barcode_{id_number}.{format}"
    my_code.save(filename)
    return filename

# Example usage:
university_id = "**********"  # Replace with your actual university ID
barcode_filename = generate_barcode(university_id, format="png")
print(f"Barcode image saved as {barcode_filename}")
