"""This module defines any initial processes to run when the robot starts."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.common import datetime_util
from robot_framework import globals
from datetime import datetime

def initialize(orchestrator_connection: OrchestratorConnection) -> None:
    """Do all custom startup initializations of the robot."""
    orchestrator_connection.log_trace("Initializing.")
    
    globals.start = datetime_util.format_datetime(datetime.today())
    print(str(globals.start))