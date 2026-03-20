import win32com.client

def get_client():

    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine

        for conn in range(application.Children.Count):
            connection = application.Children(conn)

            for sess in range(connection.Children.Count):
                session = connection.Children(sess)
                if session.Info.Transaction in ("SESSION_MANAGER", "SBWP"):
                    return session

        return None

    except Exception as e:
        print("Kunne ikke forbinde til SAP GUI:", e)
        return None