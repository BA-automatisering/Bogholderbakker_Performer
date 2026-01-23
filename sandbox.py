"""For testing: Run process.py like OpenOrchestrator would run it."""
import os
import json


from OpenOrchestrator.orchestrator_connection.connection import (
    OrchestratorConnection
)
 
from robot_framework.process import process


orchestrator_connection = OrchestratorConnection(
    "Bogholderbakker_Performer",
    os.getenv("OpenOrchestratorSQL"),
    os.getenv("OpenOrchestratorKey"),
    None,
    None
)


print("sandbox started...okay")

queue_element = orchestrator_connection.get_next_queue_element('Bogholderbakke_HåndterAfvist')
#specific_content = json.loads(queue_element.data)
#queue_framework.main()
#__main__()

process(orchestrator_connection, queue_element)
 