"""For testing: Run process.py like OpenOrchestrator would run it."""
import os
import json
import ast
from robot_framework import globals
from robot_framework import lists
from robot_framework.SQL.sql_handler import SqlHandler
#from robot_framework.SQL import sql_handler


from OpenOrchestrator.orchestrator_connection.connection import (
    OrchestratorConnection
)
 
from robot_framework.process import process
from robot_framework import queue_framework
from robot_framework import initialize
from robot_framework import reset
from datetime import date, datetime, timedelta
from OpenOrchestrator.common import datetime_util
from sqlalchemy import create_engine, text

orchestrator_connection = OrchestratorConnection(
    "Bogholderbakker_Performer_sandbox",
    os.getenv("OpenOrchestratorSQL"),
    os.getenv("OpenOrchestratorKey"),
    None,
    None,
    None
)


print("sandbox started...okay")


#specific_content = json.loads(queue_element.data)
#queue_framework.main()
#__main__()
#lists.send_manuelliste(orchestrator_connection, globals.aktuel_bogholderbakke)

#globals.start = date.today()
"""
globals.start = datetime_util.format_datetime(datetime.today())
print(str(globals.start))

sql_handler = SqlHandler(orchestrator_connection)
engine = sql_handler.get_engine()

globals.aktuel_Queue = "Bogholderbakke_HåndterAfvist"
globals.start = "22-06-2026 10:10:38"
globals.start = globals.start.split(" ")[0] + " 00:00:00"
    
queue_data_dataframe = sql_handler.get_queue_data(engine, globals.start, globals.aktuel_Queue)

for row in queue_data_dataframe.itertuples():
    print(row.Index, row.data, row.message)
    row_data = ast.literal_eval(row.data)
    #orchestrator_connection.log_info(str(row.Index)+" - "+row.data+" - "+row.message)
    globals.driftliste.append({
        "status": row.status,
        "message": row.message,
        "start_date": row.start_date,
        "created_by": row.created_by,
        "data": row.data
    })
"""
reset.open_all(orchestrator_connection)

    
n = 1
while n < 40:

    queue_element = orchestrator_connection.get_next_queue_element('Bogholderbakke_DobbeltFaktura')
    process(orchestrator_connection, queue_element)
    n += 1


    
    
    
"""    
    if not len(globals.manuelliste) == 0:
        lists.send_manuelliste(orchestrator_connection, globals.aktuel_bogholderbakke)





# -----------------------------------
import subprocess

subprocess.run(["python", "-m", "robot_framework", "pn", "cs", "ck", "args", "trigger_id"])

# uv venv
# .venv\Scripts\activate
# uv pip install -e .

"""