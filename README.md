# Game Details Updater

This script updates an Excel file with game details such as "Time to Beat", "Year", and "Score" using the HowLongToBeat API. It reads an Excel file, fetches the necessary details for each game, and updates the file accordingly.

## Table of Contents

- [Requirements](#requirements)
- [Setup](#setup)
- [Usage](#usage)
- [Configuration](#configuration)
- [Logging](#logging)

## Requirements

- Python 3.7 or higher
- The following Python packages:
  - `pandas`
  - `openpyxl`
  - `tqdm`
  - `howlongtobeatpy`

## Setup

1. **Clone the repository** or download the script to your local machine.

2. **Install the required packages**. You can use `pip` to install them:
   ```bash
   pip install pandas openpyxl tqdm howlongtobeatpy
   ```

3. **Prepare your Excel file**. Ensure the Excel file (`Games.xlsx`) is in the same directory as the script. The Excel file should have the following columns:
   - Game Title
   - Platform
   - Year
   - Genre
   - Time to Beat
   - Score
   - Status

## Usage

1. **Run the script**:
   ```bash
   python script.py
   ```

2. **Script functionality**:
   - The script reads the `Games.xlsx` file
   - It checks if the file has multiple sheets and uses the first sheet
   - It updates the columns "Time to Beat", "Year", and "Score" for rows where these columns are empty
   - "Time to Beat" values are rounded up to the nearest multiple of 0.25
   - If no details are found, the script fills the columns with `-1`

## Configuration

You can configure the script by modifying the following variables in the script:
- `file_name`: The name of the Excel file to be updated. Default is `Games.xlsx`.

## Logging

The script uses Python's `logging` module to log the progress and any issues encountered. The log messages include:
- Information about the progress of updating each game
- Warnings if no details are found for a game
- Errors if there are issues while fetching details from the API

## Example

Here is an example of the expected structure of the Excel file (`Games.xlsx`):

| Game Title | Platform | Year | Genre | Time to Beat | Score | Status |
|------------|----------|------|-------|--------------|--------|--------|
| 1944: The Loop Master | Arcade | | | | | |
| 8 Legs to Love | PC | | | | | |
| Action Fighter | Sega | | | | | |

## Notes

- Ensure that the Excel file is not open in any other program while running the script to avoid permission errors
- The script saves the updated Excel file with the same name (`Games.xlsx`)

## Troubleshooting

- If you encounter a `PermissionError`, make sure the Excel file is closed and you have the necessary permissions to read and write to the file
- If the script does not update the file as expected, check the log messages for any errors or warnings

## License

This project is licensed under the MIT License. See the **LICENSE** file for more details.
