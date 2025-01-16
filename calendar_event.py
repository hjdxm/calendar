'''
作者：周雪松
email: hjdxm@outlook.com
截至日期：2025/1/16 18:11

代码：农历生日太老火了，重要日子也容易忘记，所以创建一个脚本，生成可以放在 日历 上的活动，用来提醒我自己。
'''
from ics import Calendar, Event
from lunarcalendar import Converter, Lunar
from datetime import datetime, timedelta
import pandas as pd
import os


class Zxs_calendar():
    _demo_columns = ["event", "sdatetime", "edatetime", "isLunar", "isAllDay", "repTimes", "yearD", "monthD", "dayD", "hourD", "minuteD", "description"]

    def read_csv(self, filepath: str):
        df = pd.read_excel(filepath)
        df = df.dropna(subset=("event"))
        dateColumns = ["sdatetime", "edatetime"]
        intColumns = set(self._demo_columns).difference(set(dateColumns))
        intColumns.remove("event")
        intColumns.remove("description")
        intColumns = list(intColumns)
        for column in dateColumns:
            df[column] = pd.to_datetime(df[column], format="%Y-%m-%d-%H-%M-%S")
        df[intColumns] = df[intColumns].astype("int")
        df["description"] = df["description"].fillna(" ")
        return df

    def write_demo(self):
        filename = "zxs_calendar_demo.xlsx"
        with pd.ExcelWriter(filename) as Fl:
            pd.DataFrame(columns=self._demo_columns).to_excel(Fl, index=False)
        print(f"The demo file is writed at:{filename}")

    def add_year_month(self, date, years, months):
        year = date.year + years
        month = date.month + months
        day = date.day
        hour = date.hour
        minute = date.minute
        second = date.second
        while month > 12:
            month -= 12
            year += 1
        while month < 1:
            month += 12
            year -= 1
        return datetime(year, month, day, hour, minute, second)

    def convert_excel(self, filepath: str, write_csv: bool = True):
        df = self.read_csv(filepath)
        calendar = Calendar()
        for _, record in df.iterrows():
            for times in range(record["repTimes"] + 1):
                try:
                    sdate = self.add_year_month(record["sdatetime"], times * record["yearD"], times * record["monthD"])
                    edate = self.add_year_month(record["edatetime"], times * record["yearD"], times * record["monthD"])
                except Exception:
                    continue
                if record["isLunar"]:
                    sdate = datetime.combine(Converter.Lunar2Solar(Lunar(sdate.year, sdate.month, sdate.day)).to_date(), sdate.time())
                    edate = datetime.combine(Converter.Lunar2Solar(Lunar(edate.year, edate.month, edate.day)).to_date(), edate.time())
                for temp in [sdate, edate]:
                    temp += timedelta(hours=times * record["hourD"], minutes=times * record["minuteD"], days=times * record["dayD"])

                event = Event(name=record["event"], begin=sdate, end=edate, description=record["description"])
                if record["isAllDay"]:
                    event.make_all_day()
                calendar.events.add(event)
        if write_csv:
            filename = os.path.splitext(os.path.basename(filepath))[0] + ".ics"
            with open(filename, "w", encoding="utf-8") as Fl:
                Fl.write(calendar.serialize())
                print(f"The calendar file is write at:{filename}")


if __name__ == "__main__":
    calendar = Zxs_calendar()
    # calendar.write_demo()
    calendar.convert_excel('anniversary.xlsx')
    calendar.convert_excel('birthday.xlsx')
