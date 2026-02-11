import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Define paths
DATA_PATH = os.path.join("data", "Grad Program Exit Survey Data 2024.xlsx")
OUTPUT_DIR = "outputs"
CSV_OUTPUT = os.path.join(OUTPUT_DIR, "rankings.csv")
PLOT_OUTPUT = os.path.join(OUTPUT_DIR, "rankings_plot.png")

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

def main():
    print("Loading data...")
    # Load the dataset
    # Header=0 loads the first row as columns (Q35_1, etc.)
    df = pd.read_excel(DATA_PATH, header=0)

    # Columns containing the core course rankings
    core_cols = ['Q35_1', 'Q35_5', 'Q35_2', 'Q35_4', 'Q35_3', 'Q35_8', 'Q35_9', 'Q35_10']

    # Extract course names from the second row (index 0 in df.iloc, which corresponds to the original row 2)
    # The first row of the dataframe (index 0) contains the question text.
    course_names_row = df.iloc[0]

    col_to_name = {}
    for col in core_cols:
        full_text = course_names_row[col]
        # The text format is usually:
        # "Please place each MAcc CORE course ... - [Course Name]"
        # We want to extract the part after the last " - "
        if " - " in full_text:
            course_name = full_text.split(" - ")[-1]
        else:
            course_name = full_text # Fallback
        col_to_name[col] = course_name.strip()

    print("Identified courses:")
    for col, name in col_to_name.items():
        print(f"  {col}: {name}")

    # Extract the ranking data (rows 2 onwards in original file, so index 1 onwards in df)
    rank_data = df.iloc[1:][core_cols].copy()

    # Convert to numeric, forcing errors to NaN (just in case)
    for col in core_cols:
        rank_data[col] = pd.to_numeric(rank_data[col], errors='coerce')

    # Calculate mean rank for each course
    # Lower rank is better (1 is best)
    mean_ranks = rank_data.mean().sort_values()

    # Create a DataFrame for the results
    results = pd.DataFrame({
        'Course ID': mean_ranks.index,
        'Course Name': [col_to_name[col] for col in mean_ranks.index],
        'Mean Rank': mean_ranks.values
    })

    print("\nRanking Results:")
    print(results)

    # Save to CSV
    results.to_csv(CSV_OUTPUT, index=False)
    print(f"\nSaved rankings to {CSV_OUTPUT}")

    # Visualization
    plt.figure(figsize=(10, 6))
    sns.set_theme(style="whitegrid")

    # Create bar plot
    # We want better ranks (lower values) to appear "higher" or be more prominent?
    # Or just standard bar chart where shorter bars = better rank.
    # Let's do a standard bar chart but maybe invert the x-axis if it's horizontal?
    # Or just keep it simple: Y-axis = Mean Rank (Lower is Better)

    ax = sns.barplot(x="Mean Rank", y="Course Name", data=results, hue="Course Name", palette="viridis")

    plt.title("Average Student Ranking of MAcc Core Courses\n(Lower Score = More Beneficial)")
    plt.xlabel("Mean Rank (1 = Most Beneficial, 8 = Least Beneficial)")
    plt.ylabel("")
    plt.tight_layout()

    # Save plot
    plt.savefig(PLOT_OUTPUT)
    print(f"Saved plot to {PLOT_OUTPUT}")

if __name__ == "__main__":
    main()
