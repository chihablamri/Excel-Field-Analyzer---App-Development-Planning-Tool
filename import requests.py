import requests
import re
from collections import defaultdict
from bs4 import BeautifulSoup

def decode_secret_message(doc_url):
    """
    Takes a Google Doc URL, retrieves and parses the data, and prints a grid of characters
    forming a secret message of uppercase letters.
    
    Args:
        doc_url (str): URL of the Google Doc containing Unicode characters and coordinates
    """
    try:
        response = requests.get(doc_url)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Find the table containing the data
        tables = soup.find_all('table')
        if not tables:
            print("Debug: No tables found in the document.")
            return

        grid_data = {}
        max_x = 0
        max_y = 0

        # The data seems to be in the first table
        data_table = tables[0]
        rows = data_table.find_all('tr')

        for row in rows[1:]:  # Skip header row
            cols = row.find_all('td')
            if len(cols) == 3:
                try:
                    x = int(cols[0].text.strip())
                    char = cols[1].text.strip()
                    y = int(cols[2].text.strip())

                    if char: # Ensure character is not empty
                        grid_data[(x, y)] = char
                        max_x = max(max_x, x)
                        max_y = max(max_y, y)
                except (ValueError, IndexError):
                    continue

        if not grid_data:
            print("Debug: Could not parse coordinate data from the table.")
            return

        # Create and print the grid
        print("Secret Message:")
        print("-" * (max_x + 2))
        
        for y in range(max_y + 1):
            row = ""
            for x in range(max_x + 2):
                row += grid_data.get((x, y), " ")
            print(row)
        
        print("-" * (max_x + 2))
        
    except requests.RequestException as e:
        print(f"Error fetching document: {e}")
        print("Please ensure the document is publicly accessible and the URL is correct.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Test with the provided URL
decode_secret_message("https://docs.google.com/document/d/e/2PACX-1vTER-wL5E8YC9pxDx43gk8eIds59GtUUk4nJo_ZWagbnrH0NFvMXIw6VWFLpf5tWTZIT9P9oLIoFJ6A/pub")

def test_with_sample_data():
    """
    Test function with sample data to demonstrate the grid formation.
    """
    print("\nTesting with sample 'F' pattern:")
    
    # Sample data representing the letter 'F'
    sample_grid = {
        (0, 0): 'F', (1, 0): 'F', (2, 0): 'F',
        (0, 1): 'F',
        (0, 2): 'F', (1, 2): 'F',
        (0, 3): 'F',
        (0, 4): 'F'
    }
    
    max_x = max(pos[0] for pos in sample_grid.keys()) if sample_grid else -1
    max_y = max(pos[1] for pos in sample_grid.keys()) if sample_grid else -1
    
    print("-" * (max_x + 2))
    for y in range(max_y + 1):
        row = ""
        for x in range(max_x + 1):
            row += sample_grid.get((x, y), " ")
        print(row)
    print("-" * (max_x + 2))

# To test with sample data, uncomment the following line:
# test_with_sample_data()