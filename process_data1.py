import pandas as pd

# Define file path
file_path = r"C:\Users\mabik\Downloads\Bixie.xlsx"

# Load the Excel file
try:
    df = pd.read_excel(file_path, sheet_name='Sheet1')
except FileNotFoundError:
    print(f"Error: File not found at {file_path}")
    exit()
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit()

# Define lists of malicious and benign Event IDs
malicious_ids = [4670, 4624, 4625, 4767, 4688, 4698, 5140, 4720, 4732]
benign_ids = [4634, 4648, 4769, 4776, 4672]

# Function to assign labels
def assign_label(event_ids):
    try:
        if isinstance(event_ids, str):  # Handle potential string values
            ids = list(map(int, event_ids.split(',')))
        elif isinstance(event_ids, int):  # Handle single integer values
            ids = [event_ids]
        else:
            return 0  # Handle unexpected data types
        if any(event_id in malicious_ids for event_id in ids):
            return 1  # Malicious
        return 0  # Benign
    except Exception as e:
        print(f"Error assigning label: {e}")
        return 0  # Default to benign on error

# Apply the function to the EventId column
df['Target'] = df['EventId'].apply(assign_label)

# Save the updated data to a new Excel file
output_path = r"C:\Users\mabik\Downloads\Updated_Bixie.xlsx"
try:
    df.to_excel(output_path, index=False)
    print(f"Data updated and saved to {output_path}")
except Exception as e:
    print(f"Error saving to Excel: {e}")