import datetime, math
from decimal import *

#from .const import WEEKDAYS

# Converts operating days forward or backward one day
def convert_days(op_days, shift):
    wd = list(WEEKDAYS.keys())
    converted_days = []
    
    for i, w in enumerate(wd):
        if w in op_days:
            if shift == 1:
                converted_days.append(wd[i - 6])
            elif shift == -1:
                converted_days.append(wd[i - 1])
            else:
                converted_days.append(w)

    return sorted(converted_days, key=lambda x: wd.index(x))


# Removes half-minute and replaces post-24:00 with asterisk
def convert_time(time, fmt=None):
    if time >= 86400 and not fmt:
        time -= 86400
        asterisk = '*'
    else:
        asterisk = ''
        
    return read_time(time - (time % 60), fmt).replace(' ', '') + asterisk


# Converts booking start date to effective date (adjusting for DOTW)
def get_eff_date(date_str, day_type):

    date = datetime.datetime.strptime(date_str, '%Y-%m-%d')
    add_days = 6 - date.weekday()
    date += datetime.timedelta(days=add_days)

    if day_type == 'wk':
        monday = date + datetime.timedelta(days=1)

        # Skip Labor Day
        if monday.month == 9 and monday.day <= 7:
            return monday + datetime.timedelta(days=1)
        else:
            return monday
    elif day_type == 'sa':
        return date + datetime.timedelta(days=6)
    else:
        return date


# Determines effective date, booking, scenario to use in midday storage reports
def get_mds_header_data(mds_data, s):

    header_data = {'scenario': max([c.csc_scenario for c in mds_data[s]['csc_list']])}

    eff_dates = list(set([get_eff_date(c.csc_bk_start_date, 'wk')
                          for c in mds_data[s]['csc_list']]))
    bookings = list(set([c.csc_booking for c in mds_data[s]['csc_list']]))

    if len(eff_dates) == 1:
        header_data['eff_date'] = eff_dates[0]
    else:
        header_data['eff_date'] = None
    if len(bookings) == 1:
        header_data['booking'] = bookings[0]
    else:
        header_data['booking'] = None

    return header_data


# Converts operating days string to list
def op_days(op):
    return [w for w in WEEKDAYS if w in op]


# Calculates percentage to one decimal place
def pct(num, denom):
    getcontext().prec = 9
#    print(num)
#    print(denom)
    return Decimal(num) / Decimal(denom) * 100


# Converts int/float to string of int, adds asterisk if not whole number
def strint(x, signed=False, pct=False):
    sign = '+' if x > 0 and signed else ''

    if pct:
        dec = x.quantize(Decimal('.1'), rounding=ROUND_HALF_UP)
        return '{}{:.1f}%'.format(sign, dec)
      
    if x % 1 == 0:
        return sign + str(int(x))
    else:
        return sign + str(math.floor(x)) + '*'


# Times main function
def timer(func):
    def wrap(*args, **kwargs):
        s_time = datetime.datetime.now()
        result = func(*args, **kwargs)
        e_time = datetime.datetime.now()
        elapse = (e_time - s_time).total_seconds()
        print('\nCompleted in {:.1f} seconds'.format(elapse))
        return result
    return wrap