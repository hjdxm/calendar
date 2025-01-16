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
import argparse


class Zxs_calendar():
    _demo_columns = ["event", "sdatetime", "edatetime", "isLunar", "isAllDay", "repTimes", "yearD", "monthD", "dayD", "hourD", "minuteD", "special", "description"]

    def read_excel(self, filepath: str):
        df = pd.read_excel(filepath)
        df = df.dropna(subset=("event"))
        dateColumns = ["sdatetime", "edatetime"]
        strColumns = ["event", "special", "description"]
        intColumns = set(self._demo_columns).difference(set(dateColumns)).difference(set(strColumns))
        intColumns = list(intColumns)
        for column in dateColumns:
            df[column] = pd.to_datetime(df[column], format="%Y-%m-%d-%H-%M-%S")
        df[intColumns] = df[intColumns].astype("int")
        df["description"] = df["description"].fillna(" ")
        df[strColumns] = df[strColumns].astype("str")
        return df

    def write_template(self):
        filename = "zxs_calendar_template.xlsx"
        if os.path.exists(filename):
            print(f"The template file exists!\nPlase rename or delete it before create an new template file\nFailed template file name: {filename}")
            return
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

    def convert_excel(self, filepath: str, write_excel: bool = True):
        df = self.read_excel(filepath)
        calendar = Calendar()
        event_counter = 0
        for _, record in df.iterrows():
            special = [int(i) for i in record["special"].split(",") if i != ""]
            if -2 in special:
                continue
            for times in range(record["repTimes"] + 1):
                if -1 in special:
                    pass
                elif times not in special:
                    continue
                try:
                    sdate = self.add_year_month(record["sdatetime"], times * record["yearD"], times * record["monthD"])
                    edate = self.add_year_month(record["edatetime"], times * record["yearD"], times * record["monthD"])
                except Exception:
                    continue
                if record["isLunar"]:
                    sdate = datetime.combine(Converter.Lunar2Solar(Lunar(sdate.year, sdate.month, sdate.day)).to_date(), sdate.time())
                    edate = datetime.combine(Converter.Lunar2Solar(Lunar(edate.year, edate.month, edate.day)).to_date(), edate.time())
                sdate += timedelta(hours=times * record["hourD"], minutes=times * record["minuteD"], days=times * record["dayD"])
                edate += timedelta(hours=times * record["hourD"], minutes=times * record["minuteD"], days=times * record["dayD"])

                event = Event(name=record["event"], begin=sdate, end=edate, description=record["description"])
                if record["isAllDay"]:
                    event.make_all_day()
                calendar.events.add(event)
                event_counter += 1

        if write_excel:
            filename = os.path.splitext(os.path.basename(filepath))[0] + ".ics"
            with open(filename, "w", encoding="utf-8") as Fl:
                Fl.write(calendar.serialize())
                print(f"Then calendar total events are: {event_counter}\nThe calendar file is writed at: {filename}")

    def convert_all_dir_excel(self):
        files = os.listdir()
        files = [file for file in files if (".xlsx" in file and "~" not in file)]
        for file in files:
            self.convert_excel(file)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    calendar = Zxs_calendar()
    parser.add_argument("-t", action="store_true", help="在当前文件夹下生成 .xlsx 模板文件。")
    parser.add_argument("-a", action="store_true", help="为当前文件夹下的所有 .xlsx 文件生成对应的 .ics 文件。")
    parser.add_argument("-f", nargs="+", help="跟至少一个 .xlxs 文件，将传递的所有 .xlsx 文件转为 .ics 文件")
    args = parser.parse_args()
    if args.t:
        calendar.write_template()
    elif args.a:
        calendar.convert_all_dir_excel()
    elif args.f:
        for file in args.f:
            calendar.convert_excel(file)
    else:
        print("请传递参数，输入 --help 查看帮助。")
