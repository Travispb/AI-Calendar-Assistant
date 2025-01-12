import sys
import os
import re
import openai
import calendar
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from PyQt5.QtWidgets import (
   QApplication, QMainWindow, QLabel, QPushButton,
   QTextEdit, QVBoxLayout, QWidget, QHBoxLayout, QSplitter, QComboBox
)
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtWebEngineWidgets import QWebEngineView
from datetime import datetime, timedelta
from dateutil import parser

# Google Calendar API setup
SCOPES = ['https://www.googleapis.com/auth/calendar']

load_dotenv()

openai.api_key = os.getenv("OPENAI_API_KEY")
secrets_file_path = os.getenv("GOOGLE_CLIENT_SECRET_PATH")
flow = InstalledAppFlow.from_client_secrets_file(secrets_file_path, SCOPES)
creds = flow.run_local_server(port=0)
calendar_service = build('calendar', 'v3', credentials=creds)

# Ensure 'Calendar Assistant Calendar' exists
def get_or_create_calendar():
   """
   Ensure 'Calendar Assistant Calendar' exists, or create it if not found.
   """
   try:
       calendars = calendar_service.calendarList().list().execute().get('items', [])
       for calendar in calendars:
           if calendar['summary'] == 'Calendar Assistant Calendar':
               return calendar['id']


       # Create calendar if not found
       calendar_body = {
           'summary': 'Calendar Assistant Calendar',
           'timeZone': 'UTC'
       }
       created_calendar = calendar_service.calendars().insert(body=calendar_body).execute()
       return created_calendar['id']
   except Exception as e:
       print(f"Error creating or fetching calendar: {e}")
       return None

calendar_id = get_or_create_calendar()


def parse_relative_date(date_str):
    """
    Parses relative date expressions like 'next Friday', 'this Monday', 'tomorrow', or defaults to today's date.
    """
    today = datetime.now()
    weekdays = list(calendar.day_name)
    date_str = date_str.strip().lower()

    # Handle "tomorrow"
    if date_str == "tomorrow":
        target_date = today + timedelta(days=1)
        return target_date.strftime("%Y-%m-%d")

    # Handle specific weekdays like "this Friday" or "next Friday"
    for prefix in ["this", "next"]:
        if date_str.startswith(prefix):
            day_name = date_str[len(prefix):].strip().capitalize()
            if day_name in weekdays:
                day_index = weekdays.index(day_name)
                days_ahead = (day_index - today.weekday() + 7) % 7
                if prefix == "next" or days_ahead == 0:  # Ensure "next" moves to the next week if today matches
                    days_ahead += 7
                target_date = today + timedelta(days=days_ahead)
                return target_date.strftime("%Y-%m-%d")

    # Handle weekdays without prefixes (e.g., "Friday")
    if date_str.capitalize() in weekdays:
        day_name = date_str.capitalize()
        day_index = weekdays.index(day_name)
        days_ahead = (day_index - today.weekday() + 7) % 7
        if days_ahead == 0:  # Default to the next week if today matches
            days_ahead = 7
        target_date = today + timedelta(days=days_ahead)
        return target_date.strftime("%Y-%m-%d")

    # Default to today if no valid date is parsed
    return today.strftime("%Y-%m-%d")

def parse_recurrence(text, start_date=None, end_date=None):
    """
    Parses user input to generate an RRULE for recurring events.
    Handles weekly, monthly, yearly, and custom intervals.
    """
    recurrence_rule = ""
    freq = "WEEKLY"  # Default to weekly recurrence
    byday = []
    bymonth = None
    bymonthday = None
    interval = 1
    until = None

    try:
        print(f"Parsing recurrence from input: {text}")

        # Detect specific dates or annual recurrence (e.g., "annually on July 20")
        date_match = re.search(r"annually\s+on\s+([\w\s\d]+)", text, re.IGNORECASE)
        if date_match:
            freq = "YEARLY"
            # Parse the specific month and day from the date string
            date_parts = parser.parse(date_match.group(1), fuzzy=True)
            bymonth = date_parts.month
            bymonthday = date_parts.day
            print(f"Detected annual recurrence: Month {bymonth}, Day {bymonthday}")

        # Detect weekly or monthly patterns
        day_matches = re.findall(r"(monday|tuesday|wednesday|thursday|friday|saturday|sunday)", text, re.IGNORECASE)
        if day_matches:
            byday = [day[:2].upper() for day in day_matches]
            print(f"Detected days: {byday}")

        # Detect custom intervals (e.g., "every 2 weeks" or "every year")
        interval_match = re.search(r"every\s+(\d+)?\s*(weeks?|months?|years?)", text, re.IGNORECASE)
        if interval_match:
            interval = int(interval_match.group(1) or 1)
            freq = {
                "week": "WEEKLY",
                "month": "MONTHLY",
                "year": "YEARLY"
            }[interval_match.group(2).lower().rstrip('s')]
            print(f"Detected interval: {interval} {freq}")

        # Detect end date (if provided)
        if end_date:
            until = end_date.strftime("%Y%m%dT235959Z")
            print(f"Using specified end date: {until}")

        # Construct the RRULE
        rule = f"FREQ={freq};INTERVAL={interval}"
        if bymonth:
            rule += f";BYMONTH={bymonth}"
        if bymonthday:
            rule += f";BYMONTHDAY={bymonthday}"
        if byday:
            rule += f";BYDAY={','.join(byday)}"
        if until:
            rule += f";UNTIL={until}"

        recurrence_rule = f"RRULE:{rule}"
        print(f"Generated Recurrence Rule: {recurrence_rule}")

    except Exception as e:
        print(f"Error parsing recurrence: {e}")

    return [recurrence_rule] if recurrence_rule else []


def create_event_from_ai_output(ai_output, calendar_id=None, selected_color=None):
    """
    Parses AI output and creates a Google Calendar event.
    Handles single and recurring events with proper start and end date logic.
    """
    if not calendar_id:
        calendar_id = get_or_create_calendar()

    try:
        # Debug: Log the raw AI output
        print("Raw AI Output:\n", ai_output)

        # Parse AI output into a dictionary
        details = {}
        if isinstance(ai_output, str):
            for line in ai_output.split("\n"):
                if ": " in line:
                    key, value = line.split(": ", 1)
                    details[key.strip()] = value.strip()
        else:
            raise ValueError("Invalid AI output: Expected a string.")

        # Debug: Log the parsed details
        print("Parsed Event Details:\n", details)

        # Parse Start Date and End Date
        start_date_str = details.get("Start Date", "").strip()
        end_date_str = details.get("End Date", "").strip()
        if not start_date_str:
            raise ValueError("Start Date field is missing or empty in AI output.")
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()

        end_date = None
        if end_date_str:
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()

        # Parse Start Time
        start_time_str = details.get("Start Time", "15:00")
        start_time_parsed = parser.parse(f"{start_date} {start_time_str}", fuzzy=True)

        # Parse End Time
        end_time_str = details.get("End Time", None)
        end_time_parsed = (
            parser.parse(f"{start_date} {end_time_str}", fuzzy=True)
            if end_time_str
            else start_time_parsed + timedelta(hours=1)
        )

        # Validate time range
        if end_time_parsed <= start_time_parsed:
            raise ValueError(f"Invalid time range: Start time {start_time_parsed} is not before End time {end_time_parsed}")

        # Get the current time zone name
        user_time_zone = datetime.now().astimezone().tzname()

        # Prepare the event object
        event = {
            'summary': details.get("Title", "Untitled Event"),
            'description': details.get("Summary", "Not provided"),
            'start': {
                'dateTime': start_time_parsed.isoformat(),
                'timeZone': user_time_zone,
            },
            'end': {
                'dateTime': end_time_parsed.isoformat(),
                'timeZone': user_time_zone,
            },
            'location': details.get("Location", "Not specified"),
        }

        # Handle Recurrence
        if "yes" in details.get("Recurring", "").lower():
            recurrence_text = details["Recurring"]
            recurrence_rules = parse_recurrence(recurrence_text, start_date=start_date, end_date=end_date)
            if recurrence_rules:
                event["recurrence"] = recurrence_rules
                print("Recurring Event Created with RRULE:", recurrence_rules)

        # Add color to the event
        if selected_color:
            color_id = get_color_id(selected_color)
            if color_id:
                event["colorId"] = color_id

        # Debugging: Log the event payload
        print(f"Using Calendar ID: {calendar_id}")
        print(f"Final Event Payload: {event}")

        # Insert the event into Google Calendar
        created_event = calendar_service.events().insert(calendarId=calendar_id, body=event).execute()
        print(f"Event created: {created_event.get('htmlLink')}")
        return created_event

    except Exception as e:
        print(f"Error creating event: {e}")
        return None


def get_day_of_week(day_name):
   """
   Helper function to map day name to day of the week.
   0 = Monday, 1 = Tuesday, ..., 6 = Sunday.
   """
   days = {
       "Monday": 0,
       "Tuesday": 1,
       "Wednesday": 2,
       "Thursday": 3,
       "Friday": 4,
       "Saturday": 5,
       "Sunday": 6
   }
   return days.get(day_name, -1)

def get_color_id(color_name):
   """
   Maps user-friendly color names to Google Calendar color IDs.
   """
   color_map = {
       "Default": None, "Lavender": "1", "Sage": "2", "Grape": "3", "Flamingo": "4",
       "Banana": "5", "Tangerine": "6", "Peacock": "7", "Graphite": "8", "Blueberry": "9", "Basil": "10", "Tomato": "11"
   }
   return color_map.get(color_name, None)

def get_selected_calendar_id(selected_calendar_name):
    """
    Get the calendar ID based on the selected calendar name.
    """
    try:
        calendars = calendar_service.calendarList().list().execute().get('items', [])
        for calendar in calendars:
            if calendar['summary'] == selected_calendar_name:
                return calendar['id']
    except Exception as e:
        print(f"Error fetching calendar ID: {e}")
    return None



class CalendarApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI Calendar Assistant")
        self.setGeometry(100, 100, 1200, 700)

        # Initialize suggested event
        self.suggested_event = None

        # Main layout setup
        main_layout = QVBoxLayout()

        # Horizontal splitter for dynamic resizing
        splitter = QSplitter(Qt.Horizontal)

        # Left panel setup
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(15, 15, 15, 15)
        left_layout.setSpacing(10)

        # Text input box
        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText("Enter event details...")

        # Color selector
        self.color_selector = QComboBox()
        self.color_selector.addItems([
            "Default", "Lavender", "Sage", "Grape", "Flamingo",
            "Banana", "Tangerine", "Peacock", "Graphite", "Blueberry", "Basil", "Tomato"
        ])
        self.color_selector.setToolTip("Select Event Color")

        # Calendar selector dropdown
        self.calendar_selector = QComboBox()
        self.calendar_selector.setToolTip("Select Calendar")
        calendars = calendar_service.calendarList().list().execute().get('items', [])
        self.calendar_selector.addItem("Calendar Assistant Calendar")
        for calendar in calendars:
            self.calendar_selector.addItem(calendar['summary'])

        # Process button
        self.process_button = QPushButton("Create Event")
        self.process_button.clicked.connect(self.process_input)
        self.process_button.setStyleSheet("""
            QPushButton {
                border-radius: 8px;
                background-color: #007BFF;
                color: white;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)

        # Result label
        self.result_label = QLabel("")

        # Confirm and Reject buttons
        self.confirm_button = QPushButton("Confirm")
        self.confirm_button.setStyleSheet("""
            QPushButton {
                border-radius: 8px;
                background-color: #4CAF50;
                color: white;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #0A6A47;
            }
        """)
        self.confirm_button.clicked.connect(self.confirm_event)
        self.confirm_button.hide()

        self.reject_button = QPushButton("Reject")
        self.reject_button.setStyleSheet("""
            QPushButton {
                border-radius: 8px;
                background-color: #f44336;
                color: white;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #950606;
            }
        """)
        self.reject_button.clicked.connect(self.reject_event)
        self.reject_button.hide()

        # Dropdown layout
        dropdown_layout = QHBoxLayout()
        dropdown_layout.addWidget(self.calendar_selector)
        dropdown_layout.addWidget(self.color_selector)

        # Add widgets to left layout
        left_layout.addWidget(self.text_input)
        left_layout.addLayout(dropdown_layout)  # Add dropdowns first
        left_layout.addWidget(self.process_button)  # Create Event button below dropdowns
        left_layout.addWidget(self.result_label)
        left_layout.addWidget(self.confirm_button)
        left_layout.addWidget(self.reject_button)

        # AI Chat Input/Output
        self.chat_input = QTextEdit()
        self.chat_input.setPlaceholderText("Ask something like 'When do I have time to grocery shop?'")
        self.chat_button = QPushButton("Chat with Calendar")
        self.chat_button.setStyleSheet("""
            QPushButton {
                border-radius: 8px;
                background-color: #007BFF;
                color: white;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        self.chat_button.clicked.connect(self.chat_with_calendar)
        self.chat_output = QTextEdit()
        self.chat_output.setReadOnly(True)
        left_layout.addWidget(self.chat_input)
        left_layout.addWidget(self.chat_button)
        left_layout.addWidget(self.chat_output)

        # Left panel container
        left_panel = QWidget()
        left_panel.setLayout(left_layout)

        # Right panel setup (Google Calendar view)
        self.web_view = QWebEngineView()
        self.web_view.setUrl(QUrl("https://calendar.google.com"))

        # Splitter configuration
        splitter.addWidget(left_panel)
        splitter.addWidget(self.web_view)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([400, 800])

        # Add splitter to main layout
        main_layout.addWidget(splitter)

        # Central widget setup
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Initialize suggested event
        self.suggested_event = None


    def process_input(self):
        """
        Processes user input to create multiple tasks or events in Google Calendar.
        Ensures GPT-4 output adheres to a specific format and handles phrases like "for 6 months."
        """
        user_input = self.text_input.toPlainText()
        if not user_input.strip():
            self.result_label.setText("Input is empty. Please enter event details.")
            return

        # Get the current date for reference
        now = datetime.now()
        current_date = now.strftime("%Y-%m-%d")
        current_year = now.year

        try:
            # Pre-prompt to enforce consistent format
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[
                    {
                        "role": "system",
                        "content": (
                            f"You are a scheduling assistant. The current date is {current_date}, and the current year is {current_year}. If No start time AND no end time specified, set start to 12am, end to 12am"
                            "Always respond with events in the following format:\n\n"
                            "Event:\n"
                            "Title: <Event Title>\n"
                            "Start Date: <Start date in YYYY-MM-DD format>\n"
                            "End Date: <End date in YYYY-MM-DD format, derived from phrases like 'for 6 months' or 'until December 2025', if none specified leave blank>\n"
                            "Start Time: <Start Time in HH:MM 12-hour format, default 1 hour before End Time if not specified>\n"
                            "End Time: <End Time in HH:MM 12-hour format, default 1 hour after Start Time if not specified>\n"
                            "Summary: <Optional Summary>\n"
                            "Location: <Optional Location>\n"
                            "Recurring: <Yes/No, followed by recurrence details if Yes>\n"
                            "---\n"
                            "Separate multiple events with '---'."
                        ),
                    },
                    {"role": "user", "content": f"Classify and process: {user_input}"},
                ],
            )

            ai_output = response['choices'][0]['message']['content'].strip()

            # Split the output into lines and handle formatting
            events = ai_output.split("---")
            formatted_output = "\n\n".join(event.strip() for event in events if event.strip())

            # Display the formatted output in a readable way
            self.result_label.setText(f"Processed Output:\n\n{formatted_output}")

            for event in events:
                event = event.strip()
                if "Event:" in event:
                    self.suggested_event = event  # Pass the raw string to create_event_from_ai_output
                    self.show_next_event()

            self.text_input.clear()

        except Exception as e:
            self.result_label.setText(f"Error processing input: {e}")

    def parse_event_details(self, event_text):
        """
        Parses the event text to extract details like start date, end date, and handles phrases like "for 6 months."
        """
        details = {}
        try:
            # Parse the event text into a dictionary
            for line in event_text.split("\n"):
                if ": " in line:
                    key, value = line.split(": ", 1)
                    details[key.strip()] = value.strip()

            # Handle Start and End Dates
            start_date = details.get("Start Date", datetime.now().strftime("%Y-%m-%d"))
            end_date = details.get("End Date")

            # Calculate end date if phrases like "for 6 months" are used
            if "for" in end_date.lower():
                match = re.search(r"for\s+(\d+)\s+(days?|weeks?|months?|years?)", end_date, re.IGNORECASE)
                if match:
                    num = int(match.group(1))
                    unit = match.group(2).lower().rstrip("s")
                    time_delta = {
                        "day": timedelta(days=num),
                        "week": timedelta(weeks=num),
                        "month": timedelta(days=30 * num),  # Approximate a month as 30 days
                        "year": timedelta(days=365 * num),
                    }.get(unit, timedelta(days=0))
                    end_date_parsed = datetime.strptime(start_date, "%Y-%m-%d") + time_delta
                    end_date = end_date_parsed.strftime("%Y-%m-%d")

            # Default end date if none provided
            if not end_date:
                end_date = (datetime.strptime(start_date, "%Y-%m-%d") + timedelta(days=365)).strftime("%Y-%m-%d")

            details["Start Date"] = start_date
            details["End Date"] = end_date

            # Handle recurrence normalization
            if details.get("Recurring", "").strip().lower() == "yes" and "every" in details.get("Recurring", "").lower():
                recurrence_text = details["Recurring"]
                details["Recurring"] = parse_recurrence(recurrence_text, start_date=datetime.strptime(start_date, "%Y-%m-%d"))

            return details

        except Exception as e:
            print(f"Error parsing event details: {e}")
            return details

    def normalize_event_details(self, event_text):
        """
        Ensures the event details have a valid format, including start/end times and recurrence.
        """
        details = {}
        try:
            # Parse the event text into a dictionary
            for line in event_text.split("\n"):
                if ": " in line:
                    key, value = line.split(": ", 1)
                    details[key.strip()] = value.strip()

            # Ensure Start and End Times are properly set
            event_date = details.get("Date", datetime.now().strftime("%Y-%m-%d"))
            start_time = details.get("Start Time", "15:00")  # Default to 3:00 PM
            end_time = details.get("End Time", None)

            # Infer end time if missing
            if not end_time:
                start_time_parsed = parser.parse(f"{event_date} {start_time}", fuzzy=True)
                end_time_parsed = start_time_parsed + timedelta(hours=1)
                end_time = end_time_parsed.strftime("%H:%M")

            details["Start Time"] = start_time
            details["End Time"] = end_time

            # Normalize recurrence details
            if details.get("Recurring", "").strip().lower() == "yes" and "every" in details.get("Recurring", "").lower():
                recurrence_text = details["Recurring"]
                details["Recurring"] = parse_recurrence(recurrence_text, start_date=parser.parse(event_date))

            return details

        except Exception as e:
            print(f"Error normalizing event details: {e}")
            return details


    def show_next_event(self):
       """
       Display the suggested event for confirmation and prompt the user to confirm or reject.
       """
       if not self.suggested_event:
           self.result_label.setText("No event to display for confirmation.")
           self.confirm_button.hide()
           self.reject_button.hide()
           return

       self.result_label.setText(f"Suggested Event:\n\n{self.suggested_event}")
       self.confirm_button.show()
       self.reject_button.show()

    def confirm_event(self):
        """
        Confirms the current event and creates it in Google Calendar.
        Clears the input text box after confirmation or rejection.
        Resets dropdowns to default values.
        """
        if not self.suggested_event:
            self.result_label.setText("No event to confirm.")
            return

        # Get color and calendar options
        selected_color = self.color_selector.currentText()
        selected_calendar = self.calendar_selector.currentText()
        calendar_id = get_selected_calendar_id(selected_calendar)

        # Create the event
        created_event = create_event_from_ai_output(self.suggested_event, calendar_id=calendar_id, selected_color=selected_color)
        if created_event:
            self.result_label.setText("Event Created Successfully!")
            self.suggested_event = None
        else:
            self.result_label.setText("Failed to create the event.")

        self.confirm_button.hide()
        self.reject_button.hide()
        self.text_input.clear()  # Clear input box after confirmation

        # Reset dropdowns to default
        self.color_selector.setCurrentIndex(0)
        self.calendar_selector.setCurrentIndex(0)

    def reject_event(self):
        """
        Reject the suggested event and clear it from the queue.
        Resets dropdowns to default values.
        """
        self.result_label.setText("Event rejected.")
        self.suggested_event = None
        self.confirm_button.hide()
        self.reject_button.hide()
        self.text_input.clear()

        # Reset dropdowns to default
        self.color_selector.setCurrentIndex(0)
        self.calendar_selector.setCurrentIndex(0)


    def chat_with_calendar(self):
        """
        Interact with the AI to process user queries while considering all calendars.
        """
        user_query = self.chat_input.toPlainText()
        if not user_query.strip():
            self.chat_output.setText("Please enter a question or request.")
            return

        try:
            # Get the current date and time for context
            now = datetime.now()
            current_date = now.strftime("%Y-%m-%d")
            current_time = now.strftime("%H:%M:%S")
            current_year = now.year

            # Fetch all calendars
            calendars = calendar_service.calendarList().list().execute().get('items', [])
            calendar_names = [calendar['summary'] for calendar in calendars]

            # Fetch events from all calendars
            events = []
            for calendar in calendars:
                calendar_id = calendar['id']
                time_min = now.isoformat() + 'Z'  # 'Z' indicates UTC time
                events_result = calendar_service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min,
                    singleEvents=True,
                    orderBy='startTime'
                ).execute()
                events.extend(events_result.get('items', []))

            # Format events for AI (including start and end times if available)
            event_descriptions = [
                f"Event: {event.get('summary', 'No Title')}\n"
                f"Start: {event['start'].get('dateTime', event['start'].get('date', 'No Start Time'))}\n"
                f"End: {event['end'].get('dateTime', event['end'].get('date', 'No End Time'))}\n"
                f"Location: {event.get('location', 'No Location')}\n"
                for event in events
            ]
            formatted_events = "\n".join(event_descriptions) if event_descriptions else "No upcoming events."

            # Prepare AI query
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[
                    {
                        "role": "system",
                        "content": (
                            f"You are an intelligent calendar assistant with access to the user's calendars. Always respond in 12h time format "
                            f"The current date is {current_date}, and the current time is {current_time}. "
                            f"The user has the following calendars: {', '.join(calendar_names)}. "
                            "Current events are listed below. Provide clear and actionable responses.\n\n"
                            f"{formatted_events}"
                        )
                    },
                    {"role": "user", "content": f"User's query: {user_query}"}
                ]
            )

            # Display AI response in the chat output
            ai_response = response['choices'][0]['message']['content'].strip()
            self.chat_output.setText(ai_response)

        except Exception as e:
            self.chat_output.setText(f"Error: {e}")

    def show_next_event(self):
       """
       Display the suggested event for confirmation and prompt the user to confirm or reject.
       """
       if not self.suggested_event:
           self.result_label.setText("No event to display for confirmation.")
           self.confirm_button.hide()
           self.reject_button.hide()
           return

       self.result_label.setText(f"Suggested Event:\n\n{self.suggested_event}")
       self.confirm_button.show()
       self.reject_button.show()

if __name__ == "__main__":
   app = QApplication(sys.argv)
   main_window = CalendarApp()
   main_window.show()
   sys.exit(app.exec_())
