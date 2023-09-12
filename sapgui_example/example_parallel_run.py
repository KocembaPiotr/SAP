from sapgui import sapgui, sapguicode

SAP_MODULE = {r'Module Name': 'Module Code'}


sap = sapgui.SapGui(SAP_MODULE)
for i in range(10):
    sapgui.sap_run_new_session(sapguicode.zpp_mat, 'PKSTATUSES', option='Mass',
                               file_name=f'zpp_mat_{i}.txt')
