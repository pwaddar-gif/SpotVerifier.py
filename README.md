# 🎬 Broadcast Spot Verifier - One Click

A professional PyQt6 desktop application for broadcast operators to verify commercial spot playback against scheduled logs in seconds.

## ✨ Features

- **One-Click Verification** - Load files and verify in <2 seconds
- **Drag & Drop Support** - Simply drag CSV files onto the window
- **Smart Column Detection** - Auto-maps Creative ID, Break, Slot, Duration, Brand
- **Color-Coded Results** - 🟢 Played, 🔴 Missed, 🟡 Mismatch, 🔵 Extra
- **Real-Time Dashboard** - Summary metrics update instantly
- **Quick Filter/Search** - Filter by Creative ID, Break, Brand
- **Professional Export** - Color-coded Excel or CSV reports with timestamp
- **Dark Theme UI** - Modern broadcast-grade interface

## 📋 What It Detects

| Issue | Symbol | Description |
|-------|--------|-------------|
| Played Correctly | 🟢 | Spot aired as scheduled with correct duration |
| Missed Spots | 🔴 | Scheduled spot that never aired |
| Duration Mismatch | 🟡 | Spot aired but with wrong duration |
| Extra/Unscheduled | 🔵 | Spot that aired but wasn't scheduled |

## 🚀 Installation

### Requirements
- Python 3.8+
- Windows/Mac/Linux

### Setup

```bash
# Clone repository
git clone https://github.com/pwaddar-gif/SpotVerifier.py.git
cd SpotVerifier.py

# Install dependencies
pip install -r requirements.txt

# Run the application
python BroadcastSpotVerifier.py
```

## 📖 Usage

### Quick Start

1. **Launch the tool**
   ```bash
   python BroadcastSpotVerifier.py
   ```

2. **Load Files** (choose one method)
   - **Drag & Drop**: Drag both CSV files onto the window
   - **Browse**: Click "📂 Browse" buttons

3. **Click One Button**
   - Click `✅ VERIFY SPOTS` button
   - Results appear in <2 seconds

4. **Review Results**
   - View in color-coded tabs
   - Use search bar to filter by Creative ID, Break, or Brand
   - Summary dashboard shows counts at top

5. **Export Report**
   - `📄 Export to Excel` - Professional color-coded report
   - `📊 Export to CSV` - Data import format

### Input File Format

#### Spot Sheet (Schedule)
| Column | Example | Required |
|--------|---------|----------|
| Blaze Creative ID | CPH000030423 | ✓ |
| Break Number | 4 | ✓ |
| Slot Number | 3 | - |
| Duration | 10 (seconds) | ✓ |
| Brand | Nike | - |
| Creative Name | Nike Ad v1 | - |

#### AsRun Log (Actual Playback)
| Column | Example | Required |
|--------|---------|----------|
| Creative ID | CPH000030423 | ✓ |
| Break ID | 4 | - |
| Actual Time | 16:42:10 | - |
| Duration | 10 | ✓ |
| Delivered Time | 16:42:00 | - |
| Platform | TV | - |

**Note**: Column names are auto-detected. Tool handles variations like "Ad ID", "Com ID", "Creative", etc.

## 📊 Example Workflow

```
Original Schedule (Spot Sheet)
─────────────────────────────────
Creative ID    Break   Duration   Brand
CPH000030423   4       10s        Nike
CPH000030677   9       20s        Adidas
CPH000031000   5       15s        Puma

                    ↓ VERIFY ↓

AsRun Log (Actual Broadcast)
─────────────────────────────────
Creative ID    Break   Actual Time   Duration
CPH000030423   4       16:42:10      10s        ✅ PLAYED
CPH000031000   2       16:58:30      15s        🔵 EXTRA (unscheduled)
                                               🔴 MISSED: CPH000030677

                    ↓ REPORT ↓

Results
─────────────────────────────────
🟢 Played: 1
🔴 Missed: 1
🔵 Extra: 1
🟡 Mismatch: 0
```

## 🎯 Real-World Example: Wimbledon 2026 Stream

```
File: Wimbledon_2026-Wimbledon_2026__Day_9,_Stream_1-English-Spot-[tv]-07-07-2026.csv

1. Drag onto window → ✓ Spot Sheet Loaded
2. Drag AsRun Log → ✓ AsRun Log Loaded
3. Click [✅ VERIFY SPOTS] → Processing...
4. Results show:
   📅 Scheduled: 245
   🟢 Played: 243
   🔴 Missed: 2
   🔵 Extra: 1
   🟡 Mismatch: 0
5. Export Excel report for compliance
```

## 🛠️ Troubleshooting

### "Please select both files" error
- Ensure you've loaded **both** Spot Sheet AND AsRun Log
- Files must be in CSV format

### Column not detected
- Tool uses smart keyword matching: "Creative", "Ad ID", "Material", "Break", "Position", "Duration", "Time"
- If columns have different names, tool falls back to column position
- Add standard headers to your CSVs for best results

### No results appearing
- Check that Creative IDs match between files (case-sensitive)
- Verify duration format (numeric or string should work)
- Try searching with no filter

### Export fails
- Ensure you have write permissions to save location
- Try different folder if permission denied

## 📄 Export Format

### Excel Report (.xlsx)
- Color-coded rows:
  - 🟢 Green = Played
  - 🔴 Red = Missed
  - 🟡 Yellow = Mismatch
  - 🔵 Blue = Extra
- Professional formatting with borders and centered text
- Auto-sized columns
- Timestamp in filename

### CSV Report (.csv)
- Standard format for data import
- Columns: Status, Creative ID, Break, Slot, Duration, Brand, Air Time
- Compatible with Excel, databases, analytics tools

## 🔧 Technical Details

- **Language**: Python 3
- **UI Framework**: PyQt6
- **Data Processing**: Pandas
- **Excel Export**: OpenPyXL
- **Processing Speed**: <2 seconds for 1000+ spots
- **Memory**: Efficient; handles large broadcast schedules

## 📝 License

MIT License - Feel free to modify and distribute

## 👤 Author

Built for broadcast QC operators by Broadcast Systems

## 📞 Support

For issues or feature requests, create an issue on GitHub.

---

**🎬 Ready to verify your broadcast spots? Download, install, and run!**
