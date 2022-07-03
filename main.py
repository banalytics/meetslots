import os
import calendar
from datetime import datetime, timedelta
import pandas as pd
from gcsa.google_calendar import GoogleCalendar
from dotenv import load_dotenv


# Working hours are not accessible through the gcsa library
# TODO: Find a better way or submit a PR to gcsa
START_OF_BUSINESS = timedelta(hours=8, minutes=30)
END_OF_BUSINESS = timedelta(hours=18, minutes=30)
MEETING_DURATION = timedelta(minutes=30)


def generate_business_boundary_series(
        start_time: datetime,
        end_time: datetime,
        business_boundary_time: timedelta,
        majority_timezone: str
) -> pd.Series:
    boundary_timestamps = pd.date_range(
        start_time.replace(hour=0, minute=0, second=0, microsecond=0) + business_boundary_time,
        end_time.replace(hour=0, minute=0, second=0, microsecond=0) + business_boundary_time
    )

    boundary_timestamps = pd.Series(boundary_timestamps)

    boundary_timestamps = boundary_timestamps.dt.tz_localize(
        majority_timezone
    )

    return boundary_timestamps


load_dotenv('.env')
calendar_object = GoogleCalendar(os.getenv('EMAIL'), credentials_path='credentials.json')

start_time = datetime.utcnow()
end_time = start_time + timedelta(days=14)

upcoming_events = calendar_object.get_events(
    time_min=start_time,
    time_max=end_time,
    order_by='startTime',
    single_events=True
)

upcoming_events = pd.DataFrame([event.__dict__ for event in upcoming_events])

# We want to extract gaps between meetings, just a basic timestamp diff doesn't account for working hours, thus we will
# insert artificial events lasting from the start until the end of the business hours to simplify processing
majority_timezone = upcoming_events['timezone'].value_counts().index[0]
upcoming_events['start'] = upcoming_events['start'].dt.tz_convert(
        majority_timezone
)

upcoming_events['end'] = upcoming_events['end'].dt.tz_convert(
        majority_timezone
)

business_ends = generate_business_boundary_series(
    start_time,
    end_time,
    END_OF_BUSINESS,
    majority_timezone
)

business_starts = generate_business_boundary_series(
    start_time,
    end_time,
    START_OF_BUSINESS,
    majority_timezone
) + timedelta(days=1)

non_working_hours_events = pd.DataFrame(
    # It may seem a bit counterintuitive for why ends go to start, but note that thw non-working hours start with the
    # end of the business day and start with the beginning of the next business day
    {
        'start': business_ends,
        'end': business_starts
    }
)

upcoming_events = pd.concat(
    [
        upcoming_events,
        non_working_hours_events
    ]
)

upcoming_events.sort_values('start', inplace=True)

upcoming_events['time_to_next_meeting'] = upcoming_events['start'].shift(-1) - upcoming_events['end']

# Overlapping meetings reslt in negative values, need to cap the diff
upcoming_events['time_to_next_meeting'] = upcoming_events['time_to_next_meeting'].apply(
    lambda x: max(x, timedelta(seconds=0))
)

calendar_gaps = upcoming_events[['end', 'time_to_next_meeting']]
# Only look at gaps with a relevant length
calendar_gaps = calendar_gaps[calendar_gaps['time_to_next_meeting'] >= MEETING_DURATION]
for date in calendar_gaps['end'].dt.date.unique():
    day_of_week = calendar.day_name[date.weekday()]
    if day_of_week not in ['Saturday', 'Sunday']:
        print(day_of_week, date)
        gaps_on_date = calendar_gaps[
            calendar_gaps['end'].dt.date == date
        ]

        for gap_on_date in gaps_on_date.itertuples():
            gap_start_str = f'{gap_on_date.end.hour}:{gap_on_date.end.minute}'
            gap_end = gap_on_date.end + gap_on_date.time_to_next_meeting
            gap_end_str = datetime.strftime(gap_end, '%H:%M')
            print(datetime.strftime(gap_on_date.end, '%H:%M'), '-', gap_end_str)
