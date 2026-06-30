"""This module is the primary module of the robot framework. It collects the functionality of the rest of the framework."""

# This module is not meant to exist next to linear_framework.py in production:
# pylint: disable=duplicate-code

import sys

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueStatus
from OpenOrchestrator.common import datetime_util

from robot_framework import initialize
from robot_framework import reset
from robot_framework.exceptions import handle_error, BusinessError, log_exception
from robot_framework import process
from robot_framework import config
from robot_framework import globals
from robot_framework import lists
from robot_framework.SQL.sql_handler import SqlHandler

import ast
import json
import win32com.client
import datetime
#from datetime import datetime


def main():
    """The entry point for the framework. Should be called as the first thing when running the robot."""
    orchestrator_connection = OrchestratorConnection.create_connection_from_args()
    sys.excepthook = log_exception(orchestrator_connection)

    orchestrator_connection.log_trace("Robot Framework started.")
    initialize.initialize(orchestrator_connection)

    queue_element = None
    error_count = 0
    task_count = 0
    
    # Retry loop
    for _ in range(globals.range_max_retry_count):
        try:
            reset.reset(orchestrator_connection)

            # Queue loop
            while task_count < config.MAX_TASK_COUNT:
                task_count += 1
                
                if globals.item_count == 16 and globals.aktuel_bogholderbakke == "Bogholderbakke_HåndterAfvist":
                    reset.reset(orchestrator_connection) #Tænker at der er brug for reset inden der køres videre
                
                #queue_element = orchestrator_connection.get_next_queue_element(config.QUEUE_NAME)    
                queue_element = orchestrator_connection.get_next_queue_element(json.loads(orchestrator_connection.process_arguments)['aktuel_queue'])

                if not queue_element:
                    orchestrator_connection.log_info("Queue empty.")
                    if not len(globals.manuelliste) == 0:
                        lists.send_manuelliste(orchestrator_connection, globals.aktuel_bogholderbakke)
                    else:
                        orchestrator_connection.log_trace("manuel liste er tom")
                    
                    sql_handler = SqlHandler(orchestrator_connection)
                    engine = sql_handler.get_engine(globals.Machine_type)

                    run_date = datetime.datetime.now()
                    run_date = run_date.strftime("%d-%m-%Y")
                    
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
                     
                    break  # Break queue loop

                try:
                    #sap_gui_auto = win32com.client.GetObject("SAPGUI")
                    #orchestrator_connection.log_trace("sap_gui_auto: "+str(sap_gui_auto))
                    process.process(orchestrator_connection, queue_element)
                    orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.DONE)

                except BusinessError as error:
                    handle_error("Business Error", error, queue_element, orchestrator_connection)

            break  # Break retry loop

        # We actually want to catch all exceptions possible here.
        # pylint: disable-next = broad-exception-caught
        except Exception as error:
            error_count += 1
            handle_error(f"Process Error #{error_count}", error, queue_element, orchestrator_connection)

    
    
    
    
    reset.clean_up(orchestrator_connection)
    reset.close_all(orchestrator_connection)
    reset.kill_all(orchestrator_connection)

    if config.FAIL_ROBOT_ON_TOO_MANY_ERRORS and error_count == config.MAX_RETRY_COUNT:
        raise RuntimeError("Process failed too many times.")
