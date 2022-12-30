from copy import deepcopy
import os

from .cbp_report import cbp_report_list
from .lrr_report import lrr_report_list
from .orf_report import orf_report_list
from .rch_report import rch_report
from .rra_report import rra_report
from .vps_report import vps_report_list
from .pf_report import pf_report
from .vht_report import vht_report

reports = {'cbp': {'name': 'Bulletin',
                   'func': cbp_report_list,
                   'prev': False,
                   'subdir': 'CBP'
                  },
#           'lrr': {'name': 'Lists of Regular Runs',
#                   'func': lrr_report_list,
#                   'prev': False,
#                   'subdir': 'List of Regular Runs'
#                  },
#           'orf': {'name': 'Operator Requirements (Full-Time)',
#                   'func': orf_report_list,
#                   'prev': False,
#                   'subdir': 'Operator Requirements'
#                  },
#           'rra': {'name': 'Run Route Assignments',
#                   'func': rra_report,
#                   'prev': False,
#                   'subdir': None
#                  },
#           'vps': {'name': 'Vehicle Pullout Sheets',
#                   'func': vps_report_list,
#                   'prev': False,
#                   'subdir': 'VPS'
#                  },
#           'pf': {'name': 'Payroll Files',
#                   'func': pf_report,
#                   'prev': False,
#                   'subdir': 'Payroll Files'
#                  },
#           'rch': {'name': 'Run Characteristics Report',
#                   'func': rch_report,
#                   'prev': True,
#                   'subdir': None
#                  },
#           'vht': {'name': 'Vehicle Hours',
#                   'func': vht_report,
#                   'prev': True,
#                   'subdir': None
#                   }
#           }
#
#
## Generates inputted report
#def gen_report(rep, curr_booking, prev_booking, export_dir, gen_count):
#    if rep in reports:
#        print('    {}{}'.format('{}.'.format(gen_count).ljust(4), reports[rep]['name']))
#        
#        if reports[rep]['subdir']:
#            export_dir = os.path.join(export_dir, reports[rep]['subdir'])
#
#        cb = deepcopy(curr_booking)
#        if reports[rep]['prev']:
#            pb = deepcopy(prev_booking)
#            reports[rep]['func'](cb, pb, export_dir)
#        else:
            reports[rep]['func'](cb, export_dir)
