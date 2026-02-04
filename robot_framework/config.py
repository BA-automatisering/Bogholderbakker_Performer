"""This module contains configuration constants used across the framework"""
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import json
import os

orchestrator_connection = OrchestratorConnection(
    "Bogholderbakker_Performer",
    os.getenv("OpenOrchestratorSQL"),
    os.getenv("OpenOrchestratorKey"),
    None,
    None
)


# The number of times the robot retries on an error before terminating.
MAX_RETRY_COUNT = 3

# Whether the robot should be marked as failed if MAX_RETRY_COUNT is reached.
FAIL_ROBOT_ON_TOO_MANY_ERRORS = True

# Error screenshot config
SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
SMTP_PORT = 25
SCREENSHOT_SENDER = "robot@friend.dk"

# Constant/Credential names
ERROR_EMAIL = "Error Email Leif"


# Queue specific configs
# ----------------------

# The name of the job queue (if any)
#QUEUE_NAME = None
QUEUE_NAME = "Bogholderbakke_NulBeløb"
#QUEUE_NAME = json.loads(orchestrator_connection.process_arguments)['aktuel_queue']

# The limit on how many queue elements to process
MAX_TASK_COUNT = 100

# ----------------------
