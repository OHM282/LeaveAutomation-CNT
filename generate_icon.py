import os
from PIL import Image, ImageDraw

def create_clock_icon(output_path="clock.ico"):
    # Generate high-res image 512x512
    size = 512
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Outer circle border (Google Blue / Deep Azure #1a73e8)
    margin = 24
    draw.ellipse(
        [margin, margin, size - margin, size - margin],
        fill="#1e293b", # Dark elegant slate background ring
        outline="#2563eb", # Vibrant modern blue
        width=20
    )
    
    # Inner face (Clean crisp white/cream background)
    inner_margin = margin + 26
    draw.ellipse(
        [inner_margin, inner_margin, size - inner_margin, size - inner_margin],
        fill="#ffffff",
        outline="#e2e8f0",
        width=8
    )
    
    center = size // 2
    radius = (size - 2 * inner_margin) // 2
    
    # Draw hour ticks
    import math
    for h in range(12):
        angle = math.radians(h * 30 - 90)
        is_major = (h % 3 == 0)
        tick_len = 32 if is_major else 18
        tick_width = 10 if is_major else 5
        color = "#1e293b" if is_major else "#94a3b8"
        
        x_start = center + int((radius - 16) * math.cos(angle))
        y_start = center + int((radius - 16) * math.sin(angle))
        x_end = center + int((radius - 16 - tick_len) * math.cos(angle))
        y_end = center + int((radius - 16 - tick_len) * math.sin(angle))
        
        draw.line([x_start, y_start, x_end, y_end], fill=color, width=tick_width)
        
    # Hands: Hour hand pointing at ~7:30 (around 225 degrees)
    # 7.5 hours = 7.5 * 30 = 225 deg from top (12 o'clock = -90 deg -> 135 deg)
    h_angle = math.radians(225 - 90)
    h_len = radius * 0.52
    h_x = center + int(h_len * math.cos(h_angle))
    h_y = center + int(h_len * math.sin(h_angle))
    draw.line([center, center, h_x, h_y], fill="#1e293b", width=16)
    
    # Minute hand pointing at 30 mins (6 o'clock = 90 deg)
    m_angle = math.radians(180 - 90)
    m_len = radius * 0.72
    m_x = center + int(m_len * math.cos(m_angle))
    m_y = center + int(m_len * math.sin(m_angle))
    draw.line([center, center, m_x, m_y], fill="#f59e0b", width=12) # Pineapple Gold accent!
    
    # Second hand or Pineapple Gold dot at center
    draw.ellipse([center - 18, center - 18, center + 18, center + 18], fill="#f59e0b", outline="#b45309", width=4)
    draw.ellipse([center - 8, center - 8, center + 8, center + 8], fill="#ffffff")
    
    # Save as multi-resolution ICO
    icon_sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    img.save(output_path, format="ICO", sizes=icon_sizes)
    print(f"Saved {output_path} with sizes: {icon_sizes}")

if __name__ == "__main__":
    create_clock_icon("C:/Users/ohm_2/.gemini/antigravity/scratch/leave_automation/clock.ico")
