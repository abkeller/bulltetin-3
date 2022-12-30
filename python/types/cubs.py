from ..utils.const import WEEKDAYS, GARAGES, GAR_PLACES, DAY_TYPES
from ..utils.sort import gar_sort

class vht_data:
    def __init__(self, booking):
        ids = []
        self.blocks = []
        for b in booking.bk_blocks():
            days = ''
            for d in b.blk_operating_days:
                days += d
            id = str(b.blk_seq_no) + days
            if id not in ids:
                ids.append(id)
                self.blocks.append(b)
                
        self.garages = ['1', '7', '6', 'K', '5', 'P', 'F']
        self.garage_totals = self.vh_by_garage(self.blocks)
        #print(self.garage_totals)

    ## calculates vehicle hours by garage, including total for garage
    # returns a dictionary with keys from GARAGES
    def vh_by_garage(self, blocks):
        vh = {}
            
        for g in self.garages:
            # create a dictionary for garage
            vh[g] = {}
            
            # create a key for each day of the week in each garage dictionary
            for w in WEEKDAYS:
                vh[g][w] = 0

        for b in blocks:
            # check to see if duration should be added to operating days
            # Michael
            for w in WEEKDAYS:
                # if day in operating days list add duration to running total for each day
                if w in b.blk_operating_days:
                    vh[b.blk_garage][w] = vh[b.blk_garage][w] + b.blk_duration / 86400
                    
        # sum all days for each gagarage            
        for g in GARAGES:
            # create a list to populate with totals for each garage by day
            total = [vh[g][w] for w in WEEKDAYS]
            
            # sum the totals from the list and append 'total' key to garage list    
            vh[g]['total'] = sum(total)
        
        vh['totals'] = {}
        for w in WEEKDAYS:
            vh['totals'][w] = sum([vh[g][w] for g in GARAGES])
            
        vh['totals']['total'] = sum([vh['totals'][w] for w in WEEKDAYS])
            
        return(vh)