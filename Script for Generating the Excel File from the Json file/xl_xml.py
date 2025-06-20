import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime

def combine_datetime(date_str, time_str):
    """Combines date and time into the EPG format: YYYYMMDDhhmmss"""
    dt_str = f"{date_str} {time_str}"
    try:
        dt = pd.to_datetime(dt_str)
        return dt.strftime("%Y%m%d%H%M%S")
    except:
        return ""

def excel_to_epg_xml(excel_file, xml_file):
    # Read Excel file
    df = pd.read_excel(excel_file)

    # Create <tv> root
    tv = ET.Element("tv")

    # Iterate through each row
    for _, row in df.iterrows():
        # Construct start and stop times
        start = combine_datetime(row["start_date"], row["start_time"])
        stop = combine_datetime(row["end_date"], row["end_time"])

        # Create <programme> with attributes
        programme = ET.SubElement(tv, "programme", {
            "start": start,
            "stop": stop,
            "channel": str(row["channel"])
        })

        # Basic elements
        ET.SubElement(programme, "title").text = str(row["title"])
        ET.SubElement(programme, "sub-title").text = str(row["subtitle"])
        ET.SubElement(programme, "desc").text = str(row["description"])
        ET.SubElement(programme, "category").text = str(row["category"])
        ET.SubElement(programme, "country").text = str(row["country"])
        ET.SubElement(programme, "episode-num").text = str(row["episode_number"])
        ET.SubElement(programme, "series-name").text = str(row["series_name"])
        ET.SubElement(programme, "rating").text = str(row["rating"])
        ET.SubElement(programme, "icon").text = str(row["icon_url"])

        # Credits
        credits = ET.SubElement(programme, "credits")
        ET.SubElement(credits, "actor").text = str(row["actors"])
        ET.SubElement(credits, "director").text = str(row["director"])

    # Write to XML file
    tree = ET.ElementTree(tv)
    tree.write(xml_file, encoding="utf-8", xml_declaration=True)

    print(f"XML file '{xml_file}' generated successfully.")

# Example usage
if __name__ == "__main__":
    excel_to_epg_xml("data.xlsx", "epg_output.xml")
