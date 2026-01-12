#!/usr/bin/python

import pandas

# Time slots for display (converted to 24-hour format for sorting)
time_slots = [
    "08:00 AM",
    "08:30 AM",
    "09:00 AM",
    "09:30 AM",
    "10:00 AM",
    "10:30 AM",
    "11:00 AM",
    "11:30 AM",
    "12:00 PM",
    "12:30 PM",
    "01:00 PM",
    "01:30 PM",
    "02:00 PM",
    "02:30 PM",
    "03:00 PM",
    "03:30 PM",
    "04:00 PM",
    "04:30 PM",
    "05:00 PM",
    "05:30 PM",
    "06:00 PM",
    "06:30 PM",
    "07:00 PM",
    "07:30 PM",
    "08:00 PM",
    "08:30 PM",
    "09:00 PM",
    "09:30 PM",
]

day_order = ["M", "T", "W", "R", "F"]


def time_to_minutes(time_str):
    """Convert time string (HH:MM AM/PM) to minutes since midnight"""
    time_part, period = time_str.rsplit(" ", 1)
    hours, minutes = map(int, time_part.split(":"))
    if period == "PM" and hours != 12:
        hours += 12
    elif period == "AM" and hours == 12:
        hours = 0
    return hours * 60 + minutes


def create_weekly_schedule(csv_file, buildings_to_include=None):
    """Read parsed CSV and create weekly schedules for each building-room

    Args:
        csv_file: Path to parsed section_tally CSV
        buildings_to_include: List of building names to include. If None, includes all.
    """
    df = pandas.read_csv(csv_file)

    # Filter to specified buildings if provided
    if buildings_to_include:
        df = df[df["Bldg"].isin(buildings_to_include)]

    # Group by building and room
    grouped = df.groupby(["Bldg", "Room"])

    # Create CSV for each room
    schedules_by_room = {}
    for (bldg, room), group_df in grouped:
        schedule = create_room_schedule(group_df)
        schedules_by_room[(bldg, room)] = schedule

        # Save individual room CSV
        filename = f"schedule_{bldg.replace(' ', '_')}-{room}.csv"
        schedule.to_csv(filename)

    # Create combined schedule with rooms arranged horizontally by day
    combined_schedule = create_combined_schedule(schedules_by_room)
    combined_schedule.to_csv("room_schedules_combined.csv")

    print(f"Created room schedules for {len(grouped)} building-room combinations")
    print("Saved individual CSVs and combined schedule to room_schedules_combined.csv")


def create_room_schedule(room_df):
    """Create a weekly schedule dataframe for a single room"""
    # Initialize schedule with time slots
    schedule = pandas.DataFrame(index=time_slots)
    schedule.index.name = "Time"

    # Add columns for each day
    for day in day_order:
        schedule[day] = ""

    # Fill in the schedule
    for _, row in room_df.iterrows():
        day = row["Day"]
        if day not in day_order:
            continue

        start_time = row["Beg"]
        end_time = row["End"]
        title = row["Title"]
        prof = row["Prof"]
        type_label = row["Type"]

        # Create display text (use pipe separator instead of newlines for CSV compatibility)
        display_text = f"{title} | {prof} | {type_label}"

        # Convert times to minutes
        start_minutes = time_to_minutes(start_time)
        end_minutes = time_to_minutes(end_time)

        # Find all time slots that fall within the course duration
        course_slots = []
        for slot in time_slots:
            slot_minutes = time_to_minutes(slot)
            # Include slots that start during the course or before it ends
            if start_minutes <= slot_minutes < end_minutes:
                course_slots.append(slot)

        # Fill all slots with course info
        for slot in course_slots:
            if schedule.loc[slot, day] == "":
                schedule.loc[slot, day] = display_text

    return schedule


def create_combined_schedule(schedules_by_room):
    """Create combined schedule with rooms stacked horizontally by day"""
    # Get sorted list of all rooms
    sorted_rooms = sorted(schedules_by_room.keys())

    # Initialize combined dataframe with Time column
    combined = pandas.DataFrame(index=time_slots)
    combined.index.name = "Time"

    # For each day, add columns for each room
    for day in day_order:
        for bldg, room in sorted_rooms:
            schedule = schedules_by_room[(bldg, room)]
            col_name = f"{bldg.replace(' ', '_')}-{room}_{day}"
            combined[col_name] = schedule[day].values

    return combined


if __name__ == "__main__":
    input_csv = "section_tally_parsed.csv"
    # Filter to specific buildings (add/remove as needed)
    buildings = ["Engineering Hall", "Rowan Hall"]
    create_weekly_schedule(input_csv, buildings_to_include=buildings)
