import pandas as pd
import asyncio
import logging
from tqdm import tqdm
from howlongtobeatpy import HowLongToBeat
import math

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to fetch the time to beat, release year, and review score for a game
async def fetch_game_details(game_title):
    try:
        logging.info(f'Searching details for: {game_title}')
        hltb = HowLongToBeat()
        results_list = await hltb.async_search(game_title)
        if results_list is not None and len(results_list) > 0:
            best_element = max(results_list, key=lambda element: element.similarity)
            logging.info(f'Details found for {game_title}: {best_element.main_story} hours, {best_element.release_world}, {best_element.review_score}')
            return best_element.main_story, best_element.release_world, best_element.review_score
        else:
            logging.warning(f'No results found for {game_title}')
    except TypeError as te:
        logging.error(f"TypeError while searching details for {game_title}: {te}")
    except Exception as e:
        logging.error(f"Error while searching details for {game_title}: {e}")
    logging.warning(f'Could not find details for: {game_title}')
    return -1, -1, -1  # Return -1, -1, -1 if details are not found or an error occurs

# Wrapper to run the async function in a synchronous context
def get_game_details(game_title):
    return asyncio.run(fetch_game_details(game_title))

# Function to round up to the nearest multiple of 0.25
def round_up_to_0_25(x):
    return math.ceil(x * 4) / 4

# Main function to update the Excel file
def update_excel(file_name):
    logging.info(f'Reading the Excel file: {file_name}')
    # Read the Excel file
    xls = pd.ExcelFile(file_name, engine='openpyxl')
    sheet_name = xls.sheet_names[0]  # Use the first sheet

    df = pd.read_excel(xls, sheet_name=sheet_name)

    # Select rows where 'Time to Beat', 'Year', or 'Score' columns are empty
    rows_to_update = df[df['Time to Beat'].isna() | df['Year'].isna() | df['Score'].isna()]

    # Set up the progress bar
    progress_bar = tqdm(total=len(rows_to_update), desc='Updating game details', unit='game')

    # Iterate over the rows of the DataFrame that need to be updated
    for index, row in rows_to_update.iterrows():
        game_title = row['Game Title']
        if isinstance(game_title, str):
            time_to_beat, release_year, review_score = get_game_details(game_title)
            if pd.isna(row['Time to Beat']):
                if time_to_beat != -1:
                    time_to_beat = round_up_to_0_25(time_to_beat)
                df.at[index, 'Time to Beat'] = time_to_beat
                if time_to_beat != -1:
                    logging.info(f'Time to beat for {game_title} updated successfully')
                else:
                    logging.warning(f'Time to beat for {game_title} not found, filled with -1')
            if pd.isna(row['Year']):
                df.at[index, 'Year'] = release_year if release_year is not None else -1
                if release_year is not None:
                    logging.info(f'Release year for {game_title} updated successfully')
                else:
                    logging.warning(f'Release year for {game_title} not found, filled with -1')
            if pd.isna(row['Score']):
                df.at[index, 'Score'] = review_score if review_score is not None else -1
                if review_score is not None:
                    logging.info(f'Review score for {game_title} updated successfully')
                else:
                    logging.warning(f'Review score for {game_title} not found, filled with -1')
        else:
            logging.error(f'Invalid game title type: {type(game_title)} for {game_title}')
            if pd.isna(row['Time to Beat']):
                df.at[index, 'Time to Beat'] = -1
            if pd.isna(row['Year']):
                df.at[index, 'Year'] = -1
            if pd.isna(row['Score']):
                df.at[index, 'Score'] = -1
        progress_bar.update(1)  # Update the progress bar

    # Close the progress bar
    progress_bar.close()

    # Save the updated DataFrame back to the Excel file
    df.to_excel(file_name, index=False, engine='openpyxl')
    logging.info(f'Updated Excel file saved as: {file_name}')

# Name of the Excel file (assuming it is in the same folder as the script)
file_name = 'Games.xlsx'

# Run the update function
if __name__ == "__main__":
    logging.info('Starting the Excel update script')
    update_excel(file_name)
    logging.info('Excel update script completed')
