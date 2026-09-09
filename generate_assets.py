import zlib
import struct
import os

def create_gemini_png(size):
    width = size
    height = size
    header = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0) # 8-bit RGBA
    ihdr_chunk = b"IHDR" + ihdr
    ihdr_crc = struct.pack(">I", zlib.crc32(ihdr_chunk))
    
    raw_data = bytearray()
    cx, cy = width / 2.0, height / 2.0
    radius = width * 0.45
    
    for y in range(height):
        raw_data.append(0) # filter type 0
        for x in range(width):
            dx = x - cx
            dy = y - cy
            dist = (dx*dx + dy*dy) ** 0.5
            
            # Google Gemini blue-to-purple gradient: #1A73E8 to #8E24AA
            t = (x + y) / (width * 2.0)
            r = int(26 * (1 - t) + 142 * t)
            g = int(115 * (1 - t) + 36 * t)
            b = int(232 * (1 - t) + 170 * t)
            
            # Sparkle star effect in the center
            star_size = width * 0.28
            is_star = False
            if abs(dx) < width * 0.35 and abs(dy) < height * 0.35:
                if (abs(dx)**0.6 + abs(dy)**0.6) <= (star_size**0.6):
                    is_star = True

            if is_star:
                # White sparkle
                raw_data.extend([255, 255, 255, 255])
            elif dist <= radius:
                # Gradient badge background
                raw_data.extend([r, g, b, 255])
            elif dist <= radius + 1.2:
                # Smooth antialias edge
                alpha = int(max(0, min(255, (radius + 1.2 - dist) * 255)))
                raw_data.extend([r, g, b, alpha])
            else:
                # Transparent background
                raw_data.extend([0, 0, 0, 0])
                
    compressed = zlib.compress(bytes(raw_data))
    idat_chunk = b"IDAT" + compressed
    idat_crc = struct.pack(">I", zlib.crc32(idat_chunk))
    
    iend_chunk = b"IEND"
    iend_crc = struct.pack(">I", zlib.crc32(iend_chunk))
    
    return header + struct.pack(">I", len(ihdr)) + ihdr_chunk + ihdr_crc + \
           struct.pack(">I", len(compressed)) + idat_chunk + idat_crc + \
           struct.pack(">I", 0) + iend_chunk + iend_crc

def main():
    target_dir = os.path.join(os.path.dirname(__file__), "assets")
    os.makedirs(target_dir, exist_ok=True)
    
    sizes = [16, 32, 64, 80, 128]
    for sz in sizes:
        png_data = create_gemini_png(sz)
        file_path = os.path.join(target_dir, f"icon-{sz}.png")
        with open(file_path, "wb") as f:
            f.write(png_data)
        print(f"Generated {file_path} ({sz}x{sz}, {len(png_data)} bytes)")

if __name__ == "__main__":
    main()
