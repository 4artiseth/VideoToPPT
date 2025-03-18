# 📽️→📊 VideoToPPT

> **Turn any video into a beautiful PowerPoint presentation with a single keystroke!**

## ✨ What is VideoToPPT?

VideoToPPT is a magical screen capture tool that transforms **Coursera** videos, tutorials, and presentations into professional PowerPoint slides – instantly. Just press a key while watching your video, and the tool intelligently captures the video content, creating perfectly formatted slides.

## 🚀 Features

- **🎯 Smart Video Detection** - Automatically finds the video area on your screen, even in browser windows
- **⌨️ One-Key Capture** - Press END, F12, or 'p' to instantly take a perfectly framed screenshot
- **🖼️ Full-Screen Slides** - Each capture fills the entire slide with no margins
- **🔄 Preserves Your Edits** - Make changes to your slides between sessions without losing them

## 🏃‍♂️ Quick Start

1. **Install the requirements:**

   ```bash
   pip install -r requirements.txt
   ```

2. **Run the tool:**

   ```bash
   python video_to_ppt.py
   ```

3. **Start capturing:**

   - Position your video on screen
   - Press **END**, **F12**, or **p** to capture a slide
   - Press **ESC** to exit

4. **Find your presentation:**
   - Your slides are saved as "Introduction Module1.pptx" (by default)
   - A Temp file is automatically maintained as "Introduction Module1_TEMP.pptx" when powerpoint is open (by default)

## 🎮 Controls

| Key   | Action                           |
| ----- | -------------------------------- |
| `END` | Capture screenshot               |
| `F12` | Capture screenshot (alternative) |
| `p`   | Capture screenshot (alternative) |
| `ESC` | Exit program                     |

## 💡 Pro Tips

- **Perfect for lectures** - Capture key points without having to write notes
- **Create study guides** - Transform entire courses into visual study material
- **Build documentation** - Easily create visual guides from video tutorials
- **Keep PowerPoint open** while capturing - see your presentation grow in real-time!
- **Edit as you go** - Delete unwanted slides without disrupting the capture process

## 🛠️ Technical Details

VideoToPPT uses computer vision to detect video players on your screen. It employs:

- OpenCV for intelligent video region detection
- Pillow for high-quality screen captures
- python-pptx for PowerPoint generation
- Key debouncing to prevent accidental duplicate captures
- Automatic backup system to preserve your work

## 🔍 How It Works

1. **Video Detection**: Analyzes your screen to find the video player using aspect ratio, position, and controls detection
2. **Capture**: Takes a perfect screenshot of just the video content
3. **PowerPoint Generation**: Creates or updates a PowerPoint file with your captured slides
4. **Backup**: Maintains a single backup file for safety

## 🤔 Troubleshooting

- **Can't see the video region?** Use the manual coordinate entry when prompted
- **PowerPoint is open?** The tool creates a temporary file that you can merge later
- **Keyboard not responding?** Try an alternative capture key (END, F12, or p)

---

**Created with ❤️ for students, educators, and knowledge seekers**
