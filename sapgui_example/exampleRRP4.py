from sapgui import sapgui
from sapgui import sapguicode
import pandas as pd

SAP_MODULE = {r'Module Name': 'Module Code'}


def main():
    sapgui.sap_close()
    p40 = sapgui.SapGui(SAP_MODULE)
    df_id = pd.DataFrame({'ID': ['99999999']})
    sapgui.SAP_TMP_FILE = 'rrp4.txt'
    p40.sap_run(sapguicode.rrp4, 'var-def', '2023-01-01', '2023-01-01',
                df_id, 'lay-def')
    df_test = sapgui.sap_download_tmp_file(1, tmp_file_del=False)
    print(df_test.to_string())
    sapgui.sap_close()


if __name__ == '__main__':
    main()