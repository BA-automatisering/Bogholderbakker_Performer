"""For testing: Run process.py like OpenOrchestrator would run it."""
import os
import json
import ast
import datetime
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
#from datetime import date, datetime, timedelta
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

#globals.start = datetime_util.format_datetime(datetime.today())
"""
globals.aktuel_Queue = "Bogholderbakke_HåndterAfvist"
globals.Machine_type = "TEST"


sql_handler = SqlHandler(orchestrator_connection)
engine = sql_handler.get_engine(globals.Machine_type)

run_date = datetime.datetime.now()
run_date = run_date.strftime("%d-%m-%Y")
#run_date = "29-06-2026"

queue_data_dataframe = sql_handler.get_queue_data(engine, run_date, globals.aktuel_Queue, globals.Machine_type)

for row in queue_data_dataframe.itertuples():
    #print(row.Index, row.data, row.message)
    #row_data = ast.literal_eval(row.data)
    globals.driftliste.append({
        "status": row.status,
        "reference": row.reference,
        "message": row.message,
        "start_date": row.start_date,
        "created_by": row.created_by,
        "data": row.data
    })
    
    
lists.send_driftliste(orchestrator_connection, globals.aktuel_bogholderbakke)

"""

reset.open_all(orchestrator_connection)

    
n = 1
while n < 40:

    queue_element = orchestrator_connection.get_next_queue_element('Bogholderbakke_NulBeløb')
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