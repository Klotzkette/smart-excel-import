"""Generate simple PNG icons for the Excel add-in."""
import struct, zlib, os

def create_png(width, height, color_rgb=(15, 123, 108)):
    """Create a simple colored PNG with a document icon shape."""
    r, g, b = color_rgb
    
    # Create pixel data - simple colored square with rounded appearance
    raw_data = b''
    for y in range(height):
        raw_data += b'\x00'  # filter byte
        for x in range(width):
            # Create a simple document icon shape
            margin = width // 6
            fold = width // 4
            
            in_body = margin <= x < width - margin and margin <= y < height - margin
            in_fold = x >= width - margin - fold and y <= margin + fold and x + y >= width - fold
            
            if in_body and not in_fold:
                # Main body - primary color
                raw_data += bytes([r, g, b, 255])
            elif in_fold and x >= width - margin - fold and y >= margin:
                # Fold area - slightly lighter
                raw_data += bytes([min(r+40, 255), min(g+40, 255), min(b+40, 255), 255])
            else:
                # Transparent
                raw_data += bytes([0, 0, 0, 0])
    
    # Build PNG
    def chunk(chunk_type, data):
        c = chunk_type + data
        crc = zlib.crc32(c) & 0xffffffff
        return struct.pack('>I', len(data)) + c + struct.pack('>I', crc)
    
    png = b'\x89PNG\r\n\x1a\n'
    png += chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0))
    png += chunk(b'IDAT', zlib.compress(raw_data))
    png += chunk(b'IEND', b'')
    return png

os.makedirs('/home/user/workspace/smart-excel-import/src/assets', exist_ok=True)

for size in [16, 32, 64, 80]:
    png_data = create_png(size, size)
    with open(f'/home/user/workspace/smart-excel-import/src/assets/icon-{size}.png', 'wb') as f:
        f.write(png_data)
    print(f"Created icon-{size}.png")
