from sapgui import sapgui
from sapgui import sapguicode

SAP_MODULE = {r'Module Name': 'Module Code'}


def main():
    sapgui.sap_close()
    p40 = sapgui.SapGui(SAP_MODULE)
    sapgui.SAP_TMP_FILE = 'mb52.txt'
    p40.sap_run(sapguicode.mb52, 'PKALL')
    df_test = sapgui.sap_download_tmp_file(0, tmp_file_del=False)
    print(df_test.head(100).to_string())
    sapgui.sap_close()


if __name__ == '__main__':
    main()