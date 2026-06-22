"""This class handles all about the SQL query to the database"""
import pandas as pd
from sqlalchemy import create_engine, text
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
#from robot_framework import globals
from datetime import datetime


class SqlHandler:
    """Denne klasse får fat i databasen, og definerer sql query'en, som henter data"""
    def __init__(self, orchestrator_connection: OrchestratorConnection):
        self.orchestrator_connection = orchestrator_connection

    def get_engine(self):
        """Skaber forbindelse til OO Test db"""
        # faelles_conn_string = self.orchestrator_connection.get_constant("Conn_string_Faellessql").value
        conn_string = (
            "mssql+pyodbc://srvsqlhotel04/BAIT-DF-OO?trusted_connection=yes&driver=ODBC+Driver+17+for+SQL+Server"
        )

        engine = create_engine(conn_string)
        
        return engine

    def get_queue_data(self, engine, start, queue):
        
        query = text("""
            SELECT
                [id]
                ,[queue_name]
                ,[status]
                ,[data]
                ,[reference]
                ,[created_date]
                ,[start_date]
                ,[end_date]
                ,[message]
                ,[created_by]
            FROM [BAIT-DF-OO].[dbo].[Queues]
            WHERE queue_name = :queue AND start_date >= :start
            order by end_date desc
            """)
        start = datetime.strptime(start, "%d-%m-%Y %H:%M:%S")    
        queue_data_dataframe = pd.read_sql(
            query, 
            con=engine,
            params={"start": start, "queue": queue}
        )
        
        return queue_data_dataframe