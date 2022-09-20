import argparse
import calendar
import os

import numpy as np
import pyperclip
from datetime import datetime, timedelta
from dateutil.parser import parse
import pandas as pd
from gcsa.google_calendar import GoogleCalendar
import platform


# Working hours are not accessible through the gcsa library
# TODO: Find a better way or submit a PR to gcsa
WORK_START = timedelta(hours=8, minutes=30)
WORK_END = timedelta(hours=18, minutes=30)
DEFAULT_MEETING_DURATION = timedelta(minutes=30)


def creation_date(path_to_file):
    """
    Try to get the date that a file was created, falling back to when it was
    last modified if that isn't possible.
    See http://stackoverflow.com/a/39501288/1709587 for explanation.
    """
    if platform.system() == 'Windows':
        return os.path.getctime(path_to_file)
    else:
        stat = os.stat(path_to_file)
        try:
            return stat.st_birthtime
        except AttributeError:
            # We're probably on Linux. No easy way to get creation dates here,
            # so we'll settle for when its content was last modified.
            return stat.st_mtime


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
        if 'token.pickle' in os.listdir():
            oath_token_creation_date = datetime.utcfromtimestamp(
                # This returns a UNIX timestamp
                creation_date('token.pickle')
            )
            if (datetime.now() - datetime.utcfromtimestamp(creation_date('token.pickle'))).days > 6:
                os.remove('token.pickle')
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

    def handle_ooo_periods(self):
        # Remove gaps that fall into out of office events
        out_of_office_events = self.upcoming_events[
            self.upcoming_events['other'].dropna().apply(lambda x: x['eventType']) == 'outOfOffice'
        ]

        for index, out_of_office_event in out_of_office_events.iterrows():
            self.calendar_gaps = self.calendar_gaps[
                np.logical_not(
                    np.logical_and(
                        self.calendar_gaps['gap_start'] >= out_of_office_event['start'],
                        self.calendar_gaps['gap_start'] + self.calendar_gaps['time_to_next_meeting'] <=
                        out_of_office_event['end'])
                )
            ]

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

        # Overlapping meetings result in negative values, need to cap the diff
        self.upcoming_events['time_to_next_meeting'] = self.upcoming_events['time_to_next_meeting'].apply(
            lambda x: max(x, timedelta(seconds=0))
        )

        self.calendar_gaps = self.upcoming_events[['end', 'time_to_next_meeting']]
        self.calendar_gaps.rename(columns={'end': 'gap_start'}, inplace=True)

        self.handle_ooo_periods()

        # Only look at gaps with a relevant length
        self.calendar_gaps = self.calendar_gaps[
            self.calendar_gaps['time_to_next_meeting'] / np.timedelta64(1, "m") >= self.desired_meeting_duration
            ]

    def format_result(self):
        for date in self.calendar_gaps['gap_start'].dt.date.unique():
            day_of_week = calendar.day_name[date.weekday()]
            if day_of_week not in ['Saturday', 'Sunday']:
                self.result += f'\n{day_of_week} {date}\n'
                gaps_on_date = self.calendar_gaps[
                    self.calendar_gaps['gap_start'].dt.date == date
                ]

                for gap_on_date in gaps_on_date.itertuples():
                    gap_end = gap_on_date.gap_start + gap_on_date.time_to_next_meeting
                    gap_end_str = datetime.strftime(gap_end, '%H:%M')
                    self.result += f'{datetime.strftime(gap_on_date.gap_start, "%H:%M")}-{gap_end_str}\n'

        pyperclip.copy(self.result)

    def find_suitable_gaps(self):
        self.get_calendar_events()
        self.process_data()
        self.format_result()


def mkdatetime(datestr: str) -> datetime:
    '''
    Parses out a date from input string
    :param datestr:
    :return:
    '''
    try:
        return parse(datestr)
    except ValueError:
        raise ValueError('Incorrect Date String')


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--start-time',
                        help='Min time from when to provide scheduling slots',
                        type=mkdatetime,
                        required=False)
    parser.add_argument('--end-time',
                        help='Min time from when to provide scheduling slots',
                        type=mkdatetime,
                        required=False)
    parser.add_argument('--next-weeks',
                        help='Number of weeks from now for when to provide slots',
                        type=int,
                        required=False)
    parser.add_argument('--meeting-duration',
                        help='Expected meeting duration in minutes, defaults to 30',
                        type=int,
                        default=30,
                        required=False)
    parser.add_argument('--email',
                        help='Email from which to pull events from',
                        type=str,
                        required=True)

    args = parser.parse_args()
    args = vars(args)

    if not args['start_time'] and not args['end_time'] and not args['next_weeks']:
        raise ValueError(
            'No valid timeframe provided, please provide a start '
            'and an end time or a number of next weeks for schedulling'
        )
    elif args['start_time'] in args.keys() and args['end_time'] in args.keys() and args['next-weeks'] in args.keys():
        # Technically, we could check if these two are not the same, but with the fact that datetime has its intervals
        # but let's not deal with that for now
        raise ValueError(
            'You provided conflicting information for date ranges by a combination of '
            'start/end time and next weeks to schedule'
        )
    elif args['start_time'] and args['end_time']:
        start_time = args['start_time']
        end_time = args['end_time']
    else:
        start_time = datetime.now()
        end_time = datetime.now() + timedelta(days=7*args['next_weeks'])

    gap_finder = GapFinder(
        email=args['email'],
        start_time=start_time,
        end_time=end_time,
        desired_meeting_duration=args['meeting_duration']
    )

    gap_finder.find_suitable_gaps()
