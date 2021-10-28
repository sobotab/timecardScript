import docx
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
import time

def get_start_date() -> tuple:
    current_time=list(time.localtime())
    date=input("What is the start date to the time period? (Enter as 'mm/dd'): ")
    current_time[1]=int(date[0:2])
    current_time[2]=int(date[3:5])
    return tuple(current_time)

def get_end_date(time) ->int:
    end_date=time+(60*60*24*13)
    return end_date

def get_guard_name() -> str:
    guard_name= input("Enter guard's first and last name seperated by a space: ")
    return guard_name

def get_guard_hours() -> [int]:
    guard_hours=""
    while (len(guard_hours) != 14):
        guard_hours=input("Enter guard's hours as a comma seperated list, starting on the start date: ")
        guard_hours=guard_hours.split(',')
    return guard_hours

def get_week_days(start_date) -> [int]:
    days=[]
    for i in range(0,14):
        days.append(str(time.localtime(time.mktime(start_date)+(i*60*60*24))[1]) + "/" + \
                str(time.localtime(time.mktime(start_date)+(i*60*60*24))[2]))
    return days

def make_guard_time_card(start_date, end_date, week_days):
    time_card_template = DocxTemplate("timecard_template.docx")
    guard_name=get_guard_name()
    guard_hours=get_guard_hours()
    int_guard_hours=[float(i) for i in guard_hours]
    context={"start": str(start_date[1])+"/"+str(start_date[2]),
             "end": str(end_date[1])+"/"+str(end_date[2]),
             "week": week_days[0:7],
             "week_2": week_days[7:14],
             "name": guard_name,
             "week_hours": guard_hours[0:7],
             "week_2_hours": guard_hours[7:14],
             "week_total": str(sum(int_guard_hours[:7])),
             "week_2_total": str(sum(int_guard_hours[7:14])),
             "total_hours": str(sum(int_guard_hours))}
    time_card_template.render(context)
    return time_card_template

if __name__ == '__main__':
    all_time_cards = docx.Document()
    composer=Composer(all_time_cards)
    start_date = get_start_date()
    end_date = time.localtime(get_end_date(time.mktime(start_date)))
    week_days=get_week_days(start_date)

    keepGoing=True
    while(keepGoing):
        composer.append(make_guard_time_card(start_date, end_date, week_days))
        keepGoing = bool(input("Press 1 and enter to add a timecard or just press enter to generate a document: "))
    save_location=input("Where would you like to save this document?: ")
    composer.save(save_location)


