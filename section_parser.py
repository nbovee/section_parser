#!/usr/bin/python

import pandas
import re

display_start_time = [
    "0800",
    "0930",
    "1100",
    "1230",
    "1400",
    "1530",
    "1700",
    "1830",
    "2000",
]
display_end_time = [
    "0915",
    "1045",
    "1215",
    "1345",
    "1515",
    "1645",
    "1815",
    "1945",
    "2115",
]


def parse_section_tally(target):
    try:
        df = pandas.read_csv(target, usecols=["Title", "Instructor(s)", "Meeting Time"])
        df = df.rename(columns={"Instructor(s)": "Prof"})
    except Exception as e:
        print("Tip: Section Tally CSV file could not be read.")
        print(e)
        exit(1)

    room_col_re = r"(?P<days>[MTWRF]+)\s+(?P<Beg>\d{1,2}:\d{2}\s(?:AM|PM))\s*-\s*(?P<End>\d{1,2}:\d{2}\s(?:AM|PM))\s+(?P<Bldg>.*?)\s+(?P<Room>\d+)\s+\((?P<Type>\w+)\)"
    room_col = "Meeting Time"

    # split timeblock data - extractall handles multiple matches per row
    df1 = (
        df[room_col]
        .str.extractall(room_col_re)
        .reset_index()
        .rename(columns={"level_0": "index", "level_1": "match_num"})
    )

    # Expand multi-day entries (e.g., "TR" becomes "T" and "R" as separate rows)
    expanded_rows = []
    for _, row in df1.iterrows():
        days_str = row["days"]
        if len(days_str) > 1:
            # Multi-day entry - create one row per day
            for day in days_str:
                new_row = row.copy()
                new_row["Day"] = day
                expanded_rows.append(new_row)
        else:
            # Single day entry
            new_row = row.copy()
            new_row["Day"] = days_str
            expanded_rows.append(new_row)

    df1_parsed = pandas.DataFrame(expanded_rows)
    # Keep only the columns we need from the regex extraction
    cols_to_keep = ["index", "Beg", "End", "Bldg", "Room", "Type", "Day"]
    df1_parsed = df1_parsed[cols_to_keep]
    df1_parsed = df1_parsed.reset_index(drop=True)

    df = (
        df.drop(axis=1, columns=room_col)
        .merge(df1_parsed, left_index=True, right_on="index", how="inner")
        .drop(axis=1, columns="index")
        .reset_index(drop=True)
    )
    df["Prof"] = df["Prof"].apply(lambda x: split_and_clean_faculty(x))
    return df


def split_and_clean_faculty(prof_str):
    """Split concatenated faculty names and clean formatting

    For multiple faculty: splits on concatenated names like "Viggiano, Paul FAimino, Jacqueline Ann"
    into "Viggiano, Paul F | Aimino, Jacqueline Ann"
    """
    if not prof_str or pandas.isna(prof_str):
        return prof_str

    prof_str = str(prof_str).strip()

    # Count commas to determine how many faculty
    comma_count = prof_str.count(",")

    if comma_count <= 1:
        # Single faculty - return full name
        return prof_str

    # Multiple faculty members - split concatenated names
    # Handle various concatenation patterns:
    # 1. "Viggiano, Paul FAimino, Jacqueline Ann" - initial + lastname
    # 2. "Williams, Allison S.Oestreich, Bruce D" - period + lastname (including compound like LeMasney)
    # 3. "Singh, RajveerIslam, Md Sadman" - lowercase + lastname (only when no period)

    # Pattern 1: Capital letter (initial) followed by Capital+lowercase+comma
    result = re.sub(r"([A-Z])([A-Z][a-z]+,)", r"\1 | \2", prof_str)

    # Pattern 2: Period followed by LastName (handles compound surnames like LeMasney)
    result = re.sub(r"(\.)([A-Z][A-Za-z]*,)", r"\1 | \2", result)

    # Pattern 3: Lowercase followed by Capital+lowercase+comma
    # Only apply if there's no period (indicates no initial/formal boundary)
    if "." not in prof_str:
        result = re.sub(r"([a-z])([A-Z][a-z]+,)", r"\1 | \2", result)

    # Clean up any trailing commas and extra spaces
    result = re.sub(r",\s*$", "", result)

    return result


def save_to_csv(dataframe, filename):
    dataframe.to_csv(filename, index=False)


if __name__ == "__main__":
    section_tally_target = "section_tally.csv"
    intermediate_output_csv = "section_tally_parsed.csv"

    df = parse_section_tally(section_tally_target)
    save_to_csv(df, intermediate_output_csv)
