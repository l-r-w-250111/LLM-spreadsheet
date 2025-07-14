import subprocess
import time
import os
import uno

UNO_CONNECTION_STRING = "uno:socket,host=localhost,port=2002;urp;"

def check_libreoffice_connection():
    """
    Checks if a LibreOffice process is already running and listening on the UNO port.
    If not, it prints instructions for the user to start it manually.
    """
    try:
        # Attempt to get the component context, which will fail if LO is not running or not listening
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
        ctx = resolver.resolve(UNO_CONNECTION_STRING + "StarOffice.ComponentContext")
        print("LibreOffice is running and connected via UNO.")
        return True
    except Exception:
        print("LibreOffice is not running or not connected via UNO.")
        print("Please start LibreOffice Calc manually with the following command:")
        print(f'"C:\\Program Files\\LibreOffice\\program\\scalc.exe" --accept="{UNO_CONNECTION_STRING}" --norestore')
        return False

def stop_libreoffice(process):
    """
    This function does nothing as LibreOffice is expected to be managed manually.
    """
    pass

if __name__ == "__main__":
    # Example usage
    if check_libreoffice_connection():
        print("Connection successful. You can now interact with LibreOffice via UNO.")
    else:
        print("Connection failed. Please follow the instructions above to start LibreOffice.")