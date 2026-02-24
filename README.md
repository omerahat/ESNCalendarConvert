## esn-calendar

Convert a monthly calendar laid out in an Excel grid into a standard `.ics` calendar file that you can import into Google Calendar, Outlook, Apple Calendar, etc.

This project is tailored for ESN-style event calendars, where:

- Each day of the month appears as a number (1–31) somewhere in the grid.
- The event text for that day lives directly in the cell below the day number.
- Optional `OC:` information can be appended in the same cell (for example: `Welcome Party OC: John Doe`).

The script reads this Excel grid and generates one all-day event per populated cell in an `.ics` file.

---

### Features

- **Excel calendar grid → ICS**: Reads a headerless, month-style Excel grid (like `March Calendar.xlsx`) and produces an `.ics` file.
- **Day detection**: Scans all cells for integers between 1 and 31 and treats them as day numbers.
- **Event text mapping**: Looks at the cell directly below each detected day and uses it as the event text.
- **Next-month spillover handling**: Ignores small day numbers (e.g. 1–9) that appear after a large day (e.g. > 25) to avoid including next month’s spillover days.
- **All-day events**: Creates all-day events for a specified `year` and `month`.
- **OC field support**: If the event text contains `OC:`, the part before `OC:` becomes the event title and the text after `OC:` is placed into the event description.
- **Istanbul timezone**: Uses the `Europe/Istanbul` timezone as a base.

---

### Requirements

- **Python**: `>= 3.9`
- **Core dependencies** (from `pyproject.toml`):
  - `ics>=0.7.2`
  - `openpyxl>=3.1.5`
  - `pandas>=2.3.3`
  - `pytz>=2025.2`

You can manage dependencies with either `uv` (recommended if you already use it) or plain `pip`.

#### Install with uv

If you use `uv` and have a `pyproject.toml` / `uv.lock` in place:

```bash
uv sync
```

This will create a virtual environment and install the required packages.

#### Install with pip

Alternatively, create and activate a virtual environment, then install the dependencies manually:

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

pip install "ics>=0.7.2" "openpyxl>=3.1.5" "pandas>=2.3.3" "pytz>=2025.2"
```

---

### Expected Excel format

The script is built around a very specific layout: a single month in a calendar-style grid, stored in an `.xlsx` file (for example, `March Calendar.xlsx`).

- **No header row**: The file is read with `header=None`, so there should be no dedicated header row.
- **Days as numbers**:
  - Cells containing integer values from 1 to 31 are interpreted as days of the month.
  - Any value outside that range is ignored.
- **Events under the day cell**:
  - For a day found at `(row, column)`, the script looks at `(row + 1, column)` for event text.
  - If that cell is non-empty (and not just whitespace), it is treated as the event text.
- **Optional OC field**:
  - You can write event cells like:  
    `Welcome Party OC: John Doe & Jane Smith`
  - The script splits at `OC:`:
    - Event name: `Welcome Party`
    - Event description: `OC: John Doe & Jane Smith`
- **Noise handling**:
  - Whitespace inside the event text is normalized (multiple spaces collapsed).
  - Very short text (length < 2 characters) is ignored (useful to skip stray punctuation like a single comma).
- **Next-month spillover days**:
  - Many calendar templates show the last few days of the previous or next month.
  - The script keeps track of the largest day number seen so far.
  - If it encounters a day `< 10` after a day `> 25`, it assumes that small day belongs to the next month and **skips** it.

If your Excel file follows this pattern, the script can reliably convert it into `.ics` events.

---

### Usage

The core logic lives in `script.py`, in the function:

```python
excel_to_ics(input_file, output_file, year, month)
```

**Parameters**

- `input_file` (`str`): Path to the input Excel file (e.g. `"March Calendar.xlsx"`).
- `output_file` (`str`): Output `.ics` filename (e.g. `"ESN_Eventleri_march.ics"`).
- `year` (`int`): Calendar year for the events (e.g. `2026`).
- `month` (`int`): Calendar month (1–12) matching the Excel grid (e.g. `3` for March).

The default configuration at the bottom of `script.py` looks like this:

```python
input_excel = "March Calendar.xlsx"  # Excel file name
output_ics = "ESN_Eventleri_march.ics"

excel_to_ics(input_excel, output_ics, year=2026, month=3)
```

#### Quickstart

1. **Prepare your Excel file**
   - Place a file like `March Calendar.xlsx` in the project root, formatted as described above.
2. **Set your month and year**
   - Open `script.py` and adjust:
     - `input_excel` to your Excel filename.
     - `output_ics` to your desired `.ics` filename.
     - `year` and `month` arguments to the correct values for that file.
3. **Run the script**

   From the project directory (and with your virtual environment activated, if using one):

   ```bash
   python script.py
   ```

4. **Import the generated ICS**
   - After successful execution, an `.ics` file such as `ESN_Eventleri_march.ics` will appear in the project directory.
   - Import this file into your calendar application of choice.

---

### How it works (high level)

At a high level, the script:

1. Reads the Excel file with `pandas.read_excel(..., header=None)` into a DataFrame.
2. Iterates over each row and column:
   - Tries to cast each cell to an integer.
   - If the value is between 1 and 31, it is considered a valid day.
3. Tracks the previous day number to detect spillover days from the next month (small numbers after large ones).
4. For each valid day:
   - Reads the cell in the row immediately below, in the same column, as event text.
   - Cleans the text, skips trivial content, and optionally extracts `OC:` details.
5. Creates an `ics.Event` object:
   - `name` is the event text before `OC:` (if present).
   - `description` contains the `OC:` segment, if any.
   - `begin` is set to the passed-in `year` and `month` plus the detected day.
   - The event is marked as an **all-day** event (`make_all_day()`).
6. Adds each event into a `Calendar` object and writes the serialized result to the output `.ics` file.

The timezone is set with `pytz.timezone("Europe/Istanbul")`, which you can change in `script.py` if needed.

---

### Data flow diagram

The following Mermaid diagram summarizes the data flow:

```mermaid
flowchart LR
excelFile["ExcelCalendar(.xlsx)"]
parser["excel_to_ics()"]
events["ics.EventObjects"]
icsFile["ESN_Eventleri_month.ics"]

excelFile --> parser
parser --> events
events --> icsFile
```

---

### Limitations and assumptions

- **All-day events only**:
  - Events do not currently support specific start/end times; every event is an all-day event.
- **Single-month files**:
  - The script assumes that each Excel file represents a single month.
  - Spillover days from the next month are ignored using a simple heuristic and are not converted.
- **Manual year/month parameters**:
  - The `year` and `month` values must be provided manually when calling `excel_to_ics`.
  - They must match the month actually represented in the Excel grid.
- **Timezone fixed to Europe/Istanbul**:
  - All events are generated in the `Europe/Istanbul` timezone.
  - Change this in the script if you need a different timezone.
- **Layout-sensitive**:
  - The script relies on the “day cell + event cell directly below” pattern.
  - Significantly different layouts may not parse correctly without changes to the logic.

---

### Extending the project

Here are some ideas for extending or adapting `esn-calendar`:

- **Custom timezone**:
  - In `script.py`, change the `timezone = pytz.timezone("Europe/Istanbul")` line to your preferred timezone.
- **Non-all-day events**:
  - Instead of `make_all_day()`, you could parse specific times from the event text and set `e.begin` / `e.end` with full datetime values.
- **Different event layouts**:
  - If your Excel calendar stores events in another pattern (e.g. multiple rows per day, events in a separate sheet), adjust the scanning logic in `excel_to_ics`.
- **Additional metadata**:
  - You can parse more structured information from the event text (e.g. location, category) and populate more fields on the `Event` object.

---

### Project metadata

- **Name**: `esn-calendar`
- **Version**: `0.1.0`
- **Description**: Convert ESN-style Excel month calendars into `.ics` files.