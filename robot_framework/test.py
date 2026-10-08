"""
class BusinessError(Exception):
    pass

try:
    raise BusinessError("Fejl")
except BusinessError as e:
    print("Fanget:", e)

print("Programmet kører videre")
"""
import os
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

orchestrator_connection = OrchestratorConnection(
    "Bogholderbakker_sandbox",
    os.getenv("OpenOrchestratorSQL_prod"),
    os.getenv("OpenOrchestratorKey_prod"),
    None,
    None,
    None
)

bruger_navn='OpusBruger_Bog'
OpusUser = 'azrpa48'
password = '&5gL46IN%Yhwt5'

orchestrator_connection.log_info('TEST '+OpusUser)
orchestrator_connection.update_credential(bruger_navn, OpusUser, password)
orchestrator_connection.log_info('Password updated for '+OpusUser+ ' manuelt fra VSCode')
