from calendar import monthrange
import datetime
from dateutil.relativedelta import relativedelta

def calculate_date():

    today = datetime.date.today()
    start = today - relativedelta(day=1, months=1)
    end = datetime.date(start.year, start.month, monthrange(start.year, start.month)[1])

    return (start, end)

def make_word(count):
    n = abs(count)
    word_base = 'строк'
    last_digit = n % 10  
    if last_digit == 1:  
        return word_base + 'а'
    elif 2 <= last_digit <= 4:  
        return word_base + 'и' 
    else:  
        return word_base