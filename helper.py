import pytz # type: ignore
from datetime import datetime

class Helper():
    def __init__(self):
        self.taiwan_timezone = pytz.timezone('Asia/Taipei')

    ## Turns a raw time string of, for example: "2024-12-11 23:08:17",
    ## from taiwan timezone to the US time,
    ## then format it in the checklist format: "%Y-%m-%d / %I:%M %p" ("2024-12-11 11:08 PM")
    def convert_timezone(self, time):
        us_time = datetime.fromisoformat(time)
        us_time = us_time.astimezone(self.taiwan_timezone)
        us_time = us_time.strftime(r"%Y-%m-%d / %I:%M %p")

        return us_time