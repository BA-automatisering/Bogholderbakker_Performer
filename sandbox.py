"""For testing: Run process.py like OpenOrchestrator would run it."""
import os
import json
from robot_framework import globals
from robot_framework import lists


from OpenOrchestrator.orchestrator_connection.connection import (
    OrchestratorConnection
)
 
from robot_framework.process import process
from robot_framework import queue_framework
from robot_framework import initialize
from robot_framework import reset

orchestrator_connection = OrchestratorConnection(
    "Bogholderbakker_Performer",
    os.getenv("OpenOrchestratorSQL"),
    os.getenv("OpenOrchestratorKey"),
    None,
    None
)


print("sandbox started...okay")


#specific_content = json.loads(queue_element.data)
#queue_framework.main()
#__main__()


reset.open_all(orchestrator_connection)
queue_element = orchestrator_connection.get_next_queue_element('Bogholderbakke_DobbeltFaktura')
process(orchestrator_connection, queue_element)
if not len(globals.manuelliste) == 0:
    lists.send_manuelliste(orchestrator_connection, globals.aktuel_bogholderbakke)
    
    
n = 1
while n < 74:

    queue_element = orchestrator_connection.get_next_queue_element('Bogholderbakke_FakturaKontrolCenter')
    process(orchestrator_connection, queue_element)
    n += 1
    if not len(globals.manuelliste) == 0:
        lists.send_manuelliste(globals.aktuel_bogholderbakke)




"""
# -----------------------------------
import subprocess

subprocess.run(["python", "-m", "robot_framework", "pn", "cs", "ck", "args", "trigger_id"])

# uv venv
# .venv\Scripts\activate
# uv pip install -e .

"""