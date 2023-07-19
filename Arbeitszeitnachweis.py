import yaml
import sys
import os
import logging
from fillpdf import fillpdfs
from datetime import datetime, timedelta
from dateutil import parser
from dateutil.rrule import rrule, MONTHLY, rrulestr
from dateutil.relativedelta import relativedelta
from holidays import country_holidays
from pytimeparse.timeparse import timeparse

logging.basicConfig(format='%(levelname)s: %(message)s', level=logging.INFO)


def to_weekday(german: str) -> str:
    english = german
    english = english.replace("Montag", "MO")
    english = english.replace("Dienstag", "TU")
    english = english.replace("Mittwoch", "WE")
    english = english.replace("Donnerstag", "TH")
    english = english.replace("Freitag", "FR")
    english = english.replace("Sonnabend", "SA")
    english = english.replace("Samstag", "SA")
    english = english.replace("Sonntag", "SU")
    return english


def to_freq(german: str) -> str:
    english = german
    english = english.replace("Minütlich", "MINUTELY")
    english = english.replace("Stündlich", "HOURLY")
    english = english.replace("Täglich", "DAILY")
    english = english.replace("Wöchentlich", "WEEKLY")
    english = english.replace("Monatlich", "MONTHLY")
    english = english.replace("Yährlich", "YEARLY")
    return english


def timedelta_to_string(delta: timedelta) -> str:
    hours, remainder = divmod(delta.total_seconds(), 3600)
    minutes = remainder // 60
    string = f"{int(hours)}h"
    if int(minutes) != 0:
        string += f" {int(minutes)}m"
    return string


if len(sys.argv) < 2:
    print("No file specified")
    exit(1)

with open(sys.argv[1], "r") as stream:
    try:
        config = yaml.safe_load(stream)
    except yaml.YAMLError as exc:
        print(exc)
        exit(1)

filename = os.path.splitext(os.path.basename(sys.argv[1]))[0]
start = datetime.strptime(config["Anfang"], "%d.%m.%Y").date()
end = datetime.strptime(config["Ende"], "%d.%m.%Y").date()

os.makedirs("./" + filename, exist_ok=True)

time_ranges: list[tuple[datetime, datetime]] = list()
for entry in config["Arbeitszeiten"]:
    entry_start = max(datetime.strptime(entry["Anfang"], "%d.%m.%Y").date(), start) if "Anfang" in entry else start
    entry_end = min(datetime.strptime(entry["Ende"], "%d.%m.%Y").date(), end) if 'Ende' in entry else end
    rrule_list: [] = []
    rrule_list.append(f"FREQ={to_freq(entry['Periode'])}")
    rrule_list.append(f"UNTIL={entry_end.strftime('%Y%m%dT%f')}")
    if "Uhrzeit" in entry:
        time = parser.parse(entry["Uhrzeit"])
        rrule_list.append(f"BYMINUTE={time.minute}")
        rrule_list.append(f"BYHOUR={time.hour}")
    if "Tag" in entry:
        rrule_list.append(f"BYDAY={to_weekday(entry['Tag'])}")

    logging.info(";".join(rrule_list))
    time_ranges += [(time, time + timedelta(seconds=timeparse(entry["Dauer"]))) for time in
                    rrulestr(";".join(rrule_list), dtstart=entry_start)]

german_holidays = country_holidays('DE', subdiv='SN', language="de")
i_month = start.month
i_year = start.year
data = dict()
data["Name"] = config["Name"]["Nachname"] + ", " + config["Name"]["Vorname"]
data["Kostenstelle"] = config["Kostenstelle"]
data["Vorgesetzter"] = config["Vorgesetzter"]
data["Geburtsdatum"] = config["Geburtsdatum"]
data["Personalnummer"] = config["Personalnummer"]
data["Struktureinheit"] = config["Struktureinheit"]
data["Wochenarbeitszeit"] = config["Wochenarbeitszeit"]
data["Vertragslaufzeit"] = start.strftime("%d.%m.%Y") + " - " + end.strftime("%d.%m.%Y")

month: datetime
for month in rrule(MONTHLY, dtstart=start, until=end):
    logging.info(f"Generating pdf for {month.strftime('%B %Y')}")
    pdf_data = dict(data)
    pdf_data["Monat"] = month.month
    pdf_data["Jahr"] = month.year
    # signature date will be the 1. day of the next month
    pdf_data["Datum"] = (month + relativedelta(months=+1)).strftime("%d.%m.%Y")

    holidays_of_month = [(day, name) for day, name in
                         country_holidays('DE', subdiv='SN', years=month.year, language="de").items() if
                         day.month == month.month]
    for holiday, name in holidays_of_month:
        # Write holiday name to last column of pdf table
        pdf_data[str(holiday.day) + "_4"] = "(F) " + name

    # TODO select all working time-ranges that are inside the month
    times_of_month = [(start, end) for start, end in time_ranges if
                      start >= month and end < month + relativedelta(months=+1)]

    # compute sum of working hours for current month
    hour_sum = timedelta()
    for s, e in times_of_month:
        logging.info(f"Inserting {s.strftime('%a %d.%m.%Y')} {s.time().strftime('%H:%M')}-{e.time().strftime('%H:%M')}")
        if s not in german_holidays and e not in german_holidays:
            pdf_data[str(s.day) + "_1"] = s.time().strftime("%H:%M")
            pdf_data[str(e.day) + "_2"] = e.time().strftime("%H:%M")
        hour_sum += e - s
        pdf_data[str(s.day) + "_3"] = timedelta_to_string(e - s)

    pdf_data["Gesamtstundenzahl"] = timedelta_to_string(hour_sum)

    file = "./" + filename + "/" + month.date().strftime("%Y-%m") + ".pdf"
    fillpdfs.write_fillable_pdf("Arbeitszeitnachweis.pdf", file, pdf_data)
