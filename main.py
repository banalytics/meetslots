import calendar
import pyperclip
from datetime import datetime, timedelta
import pandas as pd
from gcsa.google_calendar import GoogleCalendar


# Working hours are not accessible through the gcsa library
# TODO: Find a better way or submit a PR to gcsa
WORK_START = timedelta(hours=8, minutes=30)
WORK_END = timedelta(hours=18, minutes=30)
DEFAULT_MEETING_DURATION = timedelta(minutes=30)


class GapFinder:
    def __init__(
            self,
            email: str,
            start_time: datetime,
            end_time: datetime,
            work_start: timedelta = WORK_START,
            work_end: timedelta = WORK_END,
            desired_meeting_duration: timedelta = DEFAULT_MEETING_DURATION,
    ):
        self.email = email
        self.start_time = start_time
        self.end_time = end_time
        self.work_start = work_start
        self.work_end = work_end
        self.majority_timezone = None
        self.desired_meeting_duration = desired_meeting_duration
        self.upcoming_events = self.get_calendar_events()
        self.calendar_gaps = pd.DataFrame()
        self.result = ''

    def generate_business_boundary_series(
            self,
            business_boundary_time: timedelta,
    ) -> pd.Series:
        boundary_timestamps = pd.date_range(
            self.start_time.replace(hour=0, minute=0, second=0, microsecond=0) + business_boundary_time,
            self.end_time.replace(hour=0, minute=0, second=0, microsecond=0) + business_boundary_time
        )

        boundary_timestamps = pd.Series(boundary_timestamps)

        boundary_timestamps = boundary_timestamps.dt.tz_localize(
            self.majority_timezone
        )

        return boundary_timestamps

    def get_calendar_events(self):
        calendar_object = GoogleCalendar(self.email, credentials_path='credentials.json')

        upcoming_events = calendar_object.get_events(
            time_min=self.start_time,
            time_max=self.end_time,
            order_by='startTime',
            single_events=True
        )

        upcoming_events = pd.DataFrame([event.__dict__ for event in upcoming_events])

        # We want to extract gaps between meetings, just a basic timestamp diff doesn't account for working hours,
        # thus we will insert artificial events lasting from the start until the end of the business hours to
        # simplify processing
        self.majority_timezone = upcoming_events['timezone'].value_counts().index[0]
        upcoming_events['start'] = upcoming_events['start'].dt.tz_convert(
                self.majority_timezone
        )

        upcoming_events['end'] = upcoming_events['end'].dt.tz_convert(
                self.majority_timezone
        )

        return upcoming_events

    def process_data(self):
        business_ends = self.generate_business_boundary_series(
            self.work_end,
        )

        business_starts = self.generate_business_boundary_series(
            self.work_start
        ) + timedelta(days=1)

        non_working_hours_events = pd.DataFrame(
            # It may seem a bit counterintuitive for why ends go to start, but note that thw non-working hours start
            # with the end of the business day and start with the beginning of the next business day
            {
                'start': business_ends,
                'end': business_starts
            }
        )

        self.upcoming_events = pd.concat(
            [
                self.upcoming_events,
                non_working_hours_events
            ]
        )

        self.upcoming_events.sort_values('start', inplace=True)

        self.upcoming_events['time_to_next_meeting'] = (
                self.upcoming_events['start'].shift(-1) -
                self.upcoming_events['end']
        )

        # Overlapping meetings reslt in negative values, need to cap the diff
        self.upcoming_events['time_to_next_meeting'] = self.upcoming_events['time_to_next_meeting'].apply(
            lambda x: max(x, timedelta(seconds=0))
        )

        self.calendar_gaps = self.upcoming_events[['end', 'time_to_next_meeting']]
        self.calendar_gaps.rename({'end': 'gap_start'}, inplace=True, axis=1)
        # Only look at gaps with a relevant length
        self.calendar_gaps = self.calendar_gaps[self.calendar_gaps['time_to_next_meeting'] >= DEFAULT_MEETING_DURATION]

    def format_result(self):
        for date in self.calendar_gaps['gap_start'].dt.date.unique():
            day_of_week = calendar.day_name[date.weekday()]
            if day_of_week not in ['Saturday', 'Sunday']:
                self.result += f'\n{day_of_week} {date}\n'
                gaps_on_date = self.calendar_gaps[
                    self.calendar_gaps['gap_start'].dt.date == date
                ]

                for gap_on_date in gaps_on_date.itertuples():
                    gap_start_str = f'{gap_on_date.gap_start.hour}:{gap_on_date.gap_start.minute}'
                    gap_end = gap_on_date.gap_start + gap_on_date.time_to_next_meeting
                    gap_end_str = datetime.strftime(gap_end, '%H:%M')
                    self.result += f'{datetime.strftime(gap_on_date.gap_start, "%H:%M")}-{gap_end_str}\n'

        pyperclip.copy(self.result)

    def find_suitable_gaps(self):
        self.get_calendar_events()
        self.process_data()
        self.format_result()

