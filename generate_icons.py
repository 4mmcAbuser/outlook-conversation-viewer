import os
import urllib.request
from PIL import Image

def generate_icons():
    assets_dir = "assets"
    if not os.path.exists(assets_dir):
        os.makedirs(assets_dir)
        
    icon_url = "https://img.icons8.com/color/512/json.png"
    temp_path = os.path.join(assets_dir, "temp_icon.png")
    
    print(f"Downloading base icon from {icon_url}...")
    try:
        # User-Agent to prevent HTTP 403 Forbidden
        req = urllib.request.Request(
            icon_url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req) as response:
            with open(temp_path, "wb") as f:
                f.write(response.read())
        print("Download complete.")
        
        # Open and resize
        with Image.open(temp_path) as img:
            sizes = [16, 32, 64, 80, 128]
            for size in sizes:
                resized_img = img.resize((size, size), Image.Resampling.LANCZOS)
                output_path = os.path.join(assets_dir, f"icon-{size}.png")
                resized_img.save(output_path, "PNG")
                print(f"Generated {output_path} ({size}x{size})")
                
        # Clean up temp
        if os.path.exists(temp_path):
            os.remove(temp_path)
            
        print("All icons successfully generated!")
        
    except Exception as e:
        print(f"Error generating icons: {e}")

if __name__ == "__main__":
    generate_icons()
