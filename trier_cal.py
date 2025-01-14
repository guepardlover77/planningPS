import re
from datetime import datetime


def sort_and_filter_events(input_file, output_file, keywords_summary, keywords_location):
    with open(input_file, 'r', encoding='utf-8') as file:
        data = file.read()

    events = re.findall(r"BEGIN:VEVENT(.*?)END:VEVENT", data, re.DOTALL)

    event_data = []
    for event in events:
        match = re.search(r"DTSTART:(\d+T\d+Z)", event)
        summary_match = re.search(r"SUMMARY:(.*?)\n", event)
        location_match = re.search(r"LOCATION:(.*?)\n", event)

        if match:
            dtstart = datetime.strptime(match.group(1), "%Y%m%dT%H%M%SZ")
            summary = summary_match.group(1).lower() if summary_match else ""
            location = location_match.group(1).lower() if location_match else ""

            if not any(keyword in summary for keyword in keywords_summary) and not any(
                    keyword in location for keyword in keywords_location):
                event_data.append((dtstart, event))

    event_data.sort(key=lambda x: x[0])

    sorted_events = "\n".join([f"BEGIN:VEVENT{event}END:VEVENT" for _, event in event_data])

    header = re.search(r"(BEGIN:VCALENDAR.*?BEGIN:VEVENT)", data, re.DOTALL).group(1)
    footer = "END:VCALENDAR"
    sorted_calendar = f"{header}\n{sorted_events}\n{footer}"

    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(sorted_calendar)


input_file = "calendrier.txt"
output_file = "filtered_sorted_calendar.txt"
keywords_summary = ["stage", "examen", "anglais"]
keywords_location = ["tp"]
sort_and_filter_events(input_file, output_file, keywords_summary, keywords_location)
