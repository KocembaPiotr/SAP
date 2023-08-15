from sapgui import sapgui
import os


def process_exists(process_name):
    process_list = os.popen('wmic process get description, processid').read()
    return True if process_name not in process_list else False


def test_sap_close():
    sapgui.sap_close()
    assert process_exists(sapgui.SAP_APP_FILE)
