"""
Generate a modern, high-resolution 'R' app icon and Windows .ico file.
"""
import os
from PIL import Image, ImageDraw, ImageFont

def generate_r_icon():
    size = (512, 512)
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 1. Background Gradient / Rounded Badge
    # We create a sleek rounded rectangle with deep navy-to-cyan gradient
    badge_margin = 24
    badge_rect = [badge_margin, badge_margin, size[0] - badge_margin, size[1] - badge_margin]
    corner_radius = 110

    # Draw smooth gradient inside rounded rectangle
    gradient_img = Image.new("RGBA", size, (0, 0, 0, 0))
    g_draw = ImageDraw.Draw(gradient_img)

    for y in range(size[1]):
        ratio = y / size[1]
        # Modern gradient: Indigo/Navy (#1E1B4B) -> Blue (#2563EB) -> Bright Cyan (#06B6D4)
        r = int(20 * (1 - ratio) + 37 * ratio)
        g = int(30 * (1 - ratio) + 99 * (1 - ratio) * 0.5 + 182 * ratio)
        b = int(120 * (1 - ratio) + 235 * ratio)
        g_draw.line([(0, y), (size[0], y)], fill=(r, g, b, 255))

    # Mask for rounded rectangle badge
    mask = Image.new("L", size, 0)
    mask_draw = ImageDraw.Draw(mask)
    mask_draw.rounded_rectangle(badge_rect, radius=corner_radius, fill=255)

    badge = Image.composite(gradient_img, img, mask)
    draw_badge = ImageDraw.Draw(badge)

    # 2. Outer glowing rim / border
    draw_badge.rounded_rectangle(
        badge_rect,
        radius=corner_radius,
        outline=(56, 189, 248, 220), # Bright Sky Blue
        width=8
    )

    # Inner subtle glow
    inner_margin = badge_margin + 12
    draw_badge.rounded_rectangle(
        [inner_margin, inner_margin, size[0] - inner_margin, size[1] - inner_margin],
        radius=corner_radius - 10,
        outline=(255, 255, 255, 40),
        width=3
    )

    # 3. Draw the Letter 'R'
    # Try finding system fonts like Segoe UI, Arial Bold, or draw a high-precision typographic R
    font = None
    font_paths = [
        r"C:\Windows\Fonts\segoeuib.ttf",  # Segoe UI Bold
        r"C:\Windows\Fonts\arialbd.ttf",   # Arial Bold
        r"C:\Windows\Fonts\calibrib.ttf"   # Calibri Bold
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                font = ImageFont.truetype(fp, 310)
                break
            except Exception:
                pass

    if font:
        # Calculate text bounding box to center it
        bbox = draw_badge.textbbox((0, 0), "R", font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        x = (size[0] - text_w) // 2 - bbox[0]
        y = (size[1] - text_h) // 2 - bbox[1] - 8

        # Shadow for 3D depth
        shadow_offset = 6
        draw_badge.text((x + shadow_offset, y + shadow_offset), "R", font=font, fill=(10, 15, 35, 180))

        # Main Bold Letter R with White/Cyan tint
        draw_badge.text((x, y), "R", font=font, fill=(255, 255, 255, 255))
    else:
        # Fallback custom geometric R drawing
        pass

    # Save PNG and multi-size ICO
    base_dir = os.path.dirname(os.path.abspath(__file__))
    png_path = os.path.join(base_dir, "app_icon.png")
    ico_path = os.path.join(base_dir, "app_icon.ico")

    badge.save(png_path, format="PNG")
    print(f"Saved PNG icon to {png_path}")

    # Generate multi-resolution ICO file (256, 128, 64, 48, 32, 16)
    icon_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    badge.save(ico_path, format="ICO", sizes=icon_sizes)
    print(f"Saved ICO icon to {ico_path}")

if __name__ == "__main__":
    generate_r_icon()
