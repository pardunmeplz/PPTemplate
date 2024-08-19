import pillow_heif
from PIL import Image
import io

def convert_heic_to_jpeg_in_memory(heic_file):
    heif_file = pillow_heif.read_heif(heic_file)
    image = Image.frombytes(
        heif_file.mode, heif_file.size, heif_file.data, "raw", heif_file.mode
    )
    image_stream = io.BytesIO()
    image.save(image_stream, format="PNG")  # You can also use "PNG"
    image_stream.seek(0)
    return image_stream