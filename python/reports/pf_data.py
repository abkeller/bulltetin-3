from ..utils.const import WEEKDAYS, GARAGES, GAR_AREA
      
# Object representing dataset for one pf report
class pf_data:
    def __init__(self, booking):
        # Get regular runs for garage; exclude box pullers
        self.duties = [d for d in booking.bk_duties() if not d.dty_pieces[0].pce_is_box_puller]
        self.duties_by_day = self.duty_data(self.duties)
        self.hours_by_day_and_garage = self.duty_by_day(self.duties_by_day)
        self.day_total = self.hours_by_day(self.hours_by_day_and_garage)
        
        ## PTO duties
        self.pto_duties = [d for d in self.duties if d.dty_type in ['PTO', 'PTO2', 'PTOF']]
        self.pto_duties_by_day = self.duty_data(self.pto_duties)
        self.pto_hours_by_day_and_garage = self.duty_by_day(self.pto_duties_by_day)
        self.pto_day_total = self.hours_by_day(self.pto_hours_by_day_and_garage)


    def duty_data(self, duties):
        duties_by_day = []
        for d in duties:
            for day in d.dty_operating_days:
                #print(d.dty_operating_days)            
                duties_by_day.append({
                    'dty_number': d.dty_number,
                    'dty_paid_time': d.dty_paid_time.days(),
                    'dty_type': d.dty_type, 
                    'dty_day': WEEKDAYS[day],
                    'dty_garage': d.dty_garage,
                    'gar_area': GAR_AREA[d.dty_garage]
                })
        return(duties_by_day)
            
            
    def duty_by_day(self, duties):
        duty_list = []
        for g in GARAGES:
            a = GAR_AREA[g]
            a = {}
            a['gar_area'] = GAR_AREA[g]
            a['garage'] = GARAGES[g]
            
            for day in WEEKDAYS:
                a[WEEKDAYS[day]] = 0
                
            a['total'] = 0
            duty_list.append(a)
            
        duty_list = sorted(duty_list, key=lambda i: i['gar_area'])
        
        for d in duties:
            for item in duty_list:
                if GARAGES[d['dty_garage']] == item['garage']:
                    item['total'] = item['total'] + d['dty_paid_time']
                    item[d['dty_day']] = item[d['dty_day']] + d['dty_paid_time']
        return (duty_list)
    
    def hours_by_day(self, duty_by_day_and_garage):
        su = 0
        m = 0
        u = 0
        w = 0
        t = 0
        f = 0
        sa = 0
        for g in duty_by_day_and_garage:
            su = su + g['Sunday']
            m = m + g['Monday']
            u = u + g['Tuesday']
            w = w + g['Wednesday']
            t = t + g['Thursday']
            f = f + g['Friday']
            sa = sa + g['Saturday']

        totals = [su, m, u, w, t, f, sa]
        total = sum(totals)
        totals.append(total)
        return totals
                    
                    
            
            
                
            