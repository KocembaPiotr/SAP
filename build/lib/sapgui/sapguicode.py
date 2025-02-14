from sapgui import sapgui
from sapgui import sapguiparam


def zmpua25(session, **kwargs) -> None:
    """
    Function to run zmpua25 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMPUA25"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById(r"wnd[0]/usr/btn%_R_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById(r"wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def rrp4(session, **kwargs) -> None:
    """
    Function to run rrp4 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, date_from: str in format %Y-%m-%d, date_to: str in format %Y-%m-%d,
                             id_list: dataframe, layout: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/sapapo/RRP4"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtSV_DTSTA").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtSV_DTEND").text = kwargs['date_to']
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSELSCR1/ssub%_SUBSCREEN_SELBLOCK:/SAPAPO/SAPLRRP_PT_ENTRY:2010/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
        try:
            session.findById("wnd[0]").sendVKey(0)
        except Exception as e:
            print(e)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    try:
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
    except Exception as e:
        print(e)
    try:
        session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarButton("ORGRID_TOOLBAR_EXPAND")
    except Exception as e:
        print(e)
    if 'layout' in kwargs:
        session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem("&LOAD")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = ""
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").setFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]/tbar[0]/btn[12]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    try:
        session.findById("wnd[1]/usr/btnBUTTON_2").press()
    except Exception as e:
        print(e)
    session.findById("wnd[0]").sendVKey(3)


def rrp1(session, **kwargs) -> None:
    """
    Function to run rrp1 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, date_from: str in format %Y-%m-%d, date_to: str in format %Y-%m-%d,
                             id_list: dataframe, layout: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/sapapo/RRP1"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtSV_DTSTA").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtSV_DTEND").text = kwargs['date_to']
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSELSCR1/ssub%_SUBSCREEN_SELBLOCK:/SAPAPO/SAPLRRP_PT_ENTRY:2010/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
        try:
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(0)
        except Exception as e:
            print(e)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]").sendVKey(8)
    try:
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
    except Exception as e:
        print(e)
    try:
        session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarButton("ORGRID_TOOLBAR_EXPAND")
    except Exception as e:
        print(e)
    if 'layout' in kwargs:
        session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem("&LOAD")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = ""
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").setFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]/tbar[0]/btn[12]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    try:
        session.findById("wnd[1]/usr/btnBUTTON_2").press()
    except Exception as e:
        print(e)
    session.findById("wnd[0]").sendVKey(3)


def zpp_mat(session, **kwargs) -> None:
    """
    Function to run zpp_mat code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zpp_mat"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def mb52(session, **kwargs) -> None:
    """
    Function to run mb52 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "mb52"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'id_list' in kwargs:
        session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def me2m(session, **kwargs) -> None:
    """
    Function to run me2m code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant:str, po_list: dataframe, layout: str
                             date_from: str in format %Y-%m-%d, date_to: str in format %Y-%m-%d,
                             file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "me2m"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_EINDT-LOW").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_EINDT-HIGH").text = kwargs['date_to']
    if 'po_list' in kwargs:
        session.findById("wnd[0]/usr/btn%_S_EBELN_%_APP_%-VALU_PUSH").press()
        kwargs['po_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    if 'layout' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = ""
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").setFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]/tbar[0]/btn[12]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def m_ld(session, **kwargs) -> None:
    """
        Function to run m_ld code
        :param session: parameter obtained from sapgui
        :param kwargs: optional: id_list: dataframe,
                                 date_from: str in format %Y-%m-%d, date_to: str in format %Y-%m-%d,
                                 file_path, file_name : str
        :return: None
        """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "M_LD"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtRV14A-KONLI").text = "ZE"
    session.findById("wnd[0]/usr/ctxtRV14A-KONLI").caretPosition = 2
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/usr/ctxtL_3-LOW").text = "PLP2"
    session.findById("wnd[0]/usr/ctxtL_3-LOW").SetFocus()
    session.findById("wnd[0]/usr/ctxtL_3-LOW").caretPosition = 4
    session.findById("wnd[0]/usr/btn%_L_1_%_APP_%-VALU_PUSH").press()
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtDATUM-LOW").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtDATUM-HIGH").text = kwargs['date_to']
    session.findById("wnd[0]/usr/chkPAR_DAT").Selected = True
    session.findById("wnd[0]/usr/chkPAR_STAF").Selected = True
    session.findById("wnd[0]/usr/txtMAX_LINE").text = "999999"
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def mb51(session, **kwargs) -> None:
    """
    Function to run MB51 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, batch_list: dataframe,
                             date_from: str in format %Y-%m-%d, date_to: str in format %Y-%m-%d,
                             file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'id_list' in kwargs:
        session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = kwargs['date_to']
    if 'batch_list' in kwargs:
        session.findById("wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH").press()
        kwargs['batch_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zmmla11(session, **kwargs) -> None:
    """
    Function to run zmmla11 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zmmla11"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'id_list' in kwargs:
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zka02(session, **kwargs) -> None:
    """
    Function to run zka02 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: plant: str, variant: str, id_list: dataframe, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZKA02"
    session.findById("wnd[0]").sendVKey(0)
    if 'plant' in kwargs:
        session.findById("wnd[0]/usr/ctxtP_WERKS").text = kwargs['plant']
    session.findById("wnd[0]/usr/btn%P200010_1000").press()
    if 'variant' in kwargs:
        session.findById("wnd[1]/tbar[0]/btn[17]").press()
        session.findById("wnd[2]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[2]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[2]/tbar[0]/btn[8]").press()
    if 'plant' in kwargs:
        session.findById("wnd[1]/usr/ctxtW_WERKS-LOW").text = kwargs['plant']
    if 'id_list' in kwargs:
        session.findById("wnd[1]/usr/btn%_W_MATNR_%_APP_%-VALU_PUSH").press()
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[2]/tbar[0]/btn[16]").press()
        session.findById("wnd[2]/tbar[0]/btn[24]").press()
        session.findById("wnd[2]/tbar[0]/btn[8]").press()
    try:
        session.findById("wnd[1]/usr/btnIC_MORE1").press()
        if 'costing_date' in kwargs:
            session.findById("wnd[1]/usr/ctxtW_AMDAT").text = kwargs['costing_date']
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        try:
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
        except Exception as e:
            print(e)
        session.findById("wnd[1]/usr/cntlCKKK_0200_CUSTOM_CTRL/shellcont/shell").selectAll()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        if 'file_path' in kwargs:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
        else:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
        if 'file_name' in kwargs:
            sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
        else:
            sapgui.sap_del_tmp_file()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
        session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
        session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
        session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)
    except Exception as e:
        print(e)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(3)


def coois(session, **kwargs) -> None:
    """
    Function to run COOIS code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, order_list: dataframe,
                   date_from: str in format %Y-%m-%d, date_to: str in format %Y-%m-%d,
                   date_act_finish_from: str in format %Y-%m-%d, date_act_finish_to: str in format %Y-%m-%d,
                   file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "COOIS"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'order_list' in kwargs:
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press()
        kwargs['order_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").text = kwargs['date_to']
    if 'date_act_finish_from' in kwargs:
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTEN-LOW").text = kwargs['date_act_finish_from']
    if 'date_act_finish_to' in kwargs:
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ISTEN-HIGH").text = kwargs['date_act_finish_to']
    session.findById("wnd[0]").sendVKey(8)
    try:
        session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")
    except Exception as e:
        print(e)
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zcp04(session, **kwargs) -> None:
    """
    Function to run ZCP04 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe,
                   date_from: str in format %Y-%m-%d, date_to: str in format %Y-%m-%d,
                   plant: str, layout: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zcp04"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'id_list' in kwargs:
        session.findById("wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press()
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_GLTRP-LOW").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_GLTRP-HIGH").text = kwargs['date_to']
    if 'plant' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = kwargs['plant']
    else:
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = 'PLP2'
    session.findById("wnd[0]").sendVKey(8)
    if 'layout' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").Selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").SetFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]/tbar[0]/btn[12]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def me5a(session, **kwargs) -> None:
    """
    Function to run me5a code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, layout: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "me5a"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'id_list' in kwargs:
        session.findById("wnd[0]/usr/btn%_BA_BANFN_%_APP_%-VALU_PUSH").press()
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    if 'layout' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").Selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").SetFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]/tbar[0]/btn[12]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def rrp7(session, **kwargs) -> None:
    """
    Function to run rrp7 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, offset_value: str,
                             date_from: str in format %Y-%m-%d, date_to: str in format %Y-%m-%d,
                             layout: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/sapapo/RRP7"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'offset_value' in kwargs:
        session.findById("wnd[0]/usr/txtSV_ERHOF").text = kwargs['offset_value']
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtSO_START-LOW").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtSO_START-HIGH").text = kwargs['date_to']
    session.findById("wnd[0]").sendVKey(8)
    try:
        session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarButton("ORGRID_TOOLBAR_EXPAND")
    except Exception as e:
        print(e)
    if 'layout' in kwargs:
        session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem("&LOAD")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = ""
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").setFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]/tbar[0]/btn[12]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    try:
        session.findById("wnd[1]/usr/btnBUTTON_2").press()
    except Exception as e:
        print(e)
    session.findById("wnd[0]").sendVKey(3)


def me80fn(session, **kwargs) -> None:
    """
    Function to run me80fn code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, layout: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ME80FN"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'id_list' in kwargs:
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    try:
        if 'layout' in kwargs:
            session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
            session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").selectContextMenuItem("&LOAD")
            session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
            session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
            session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
            session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
            session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").Selected = True
            session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
            session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").SetFocus()
            session.findById("wnd[2]/tbar[0]/btn[0]").press()
            session.findById("wnd[2]/tbar[0]/btn[12]").press()
            session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
        session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").selectContextMenuItem("&PC")
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        if 'file_path' in kwargs:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
        else:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
        if 'file_name' in kwargs:
            sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
        else:
            sapgui.sap_del_tmp_file()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
        session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]").sendVKey(3)
        try:
            session.findById("wnd[1]/usr/btnBUTTON_2").press()
        except Exception as e:
            print(e)
    except Exception as e:
        print(e)
    session.findById("wnd[0]").sendVKey(3)


def zpoedi(session, **kwargs) -> None:
    """
    Function to run zpoedi code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZPOEDI"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'id_list' in kwargs:
        session.findById("wnd[0]/usr/ctxtEN_EBELN-LOW").text = "1"
        session.findById("wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH").press()
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    try:
        session.findById("wnd[1]/usr/btnBUTTON_2").press()
    except Exception as e:
        print(e)
    session.findById("wnd[0]").sendVKey(3)


def zca07(session, **kwargs) -> None:
    """
    Function to run zca07 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: plant: str, date_from: str in format %Y-%m-%d, date_to: str in format %Y-%m-%d,
                   id_list: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zca07"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ''
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'plant' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = kwargs['plant']
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_DATUV-LOW").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_DATUV-HIGH").text = kwargs['date_to']
    if 'id_list' in kwargs:
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/radP_CHECK8").Select()
    session.findById("wnd[0]/usr/radP_CHECK8").SetFocus()
    session.findById("wnd[0]/usr/txtP_DAYS").text = "999"
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    try:
        session.findById("wnd[1]/usr/btnBUTTON_2").press()
    except Exception as e:
        print(e)
    session.findById("wnd[0]").sendVKey(3)


def zmima18(session, **kwargs) -> None:
    """
    Function to run zmima18 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMIMA18"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def ppl1(session, **kwargs) -> None:
    """
    Function to run ppl1 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, date_from: str in format %y-%m-%d, date_to: str in format %y-%m-%d,
                             layout: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/sapapo/ppl1"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/txtSV_DTSTA").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/txtSV_DTEND").text = kwargs['date_to']
    session.findById("wnd[0]").sendVKey(8)
    if 'layout' in kwargs:
        session.findById("wnd[0]/usr/cntlCONTAIN1/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        session.findById("wnd[0]/usr/cntlCONTAIN1/shellcont/shell").selectContextMenuItem("&LOAD")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").Selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").SetFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]/tbar[0]/btn[12]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/usr/cntlCONTAIN1/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTAIN1/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zcs11h(session, **kwargs) -> None:
    """
    Function to run zcs11h code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, valid_date: str in format %y-%m-%d, id_list: dataframe,
                             file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZCS11H"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'valid_date' in kwargs:
        session.findById("wnd[0]/usr/ctxtP_DATUV").text = kwargs['valid_date']
        session.findById("wnd[0]/usr/ctxtP_DATUV").SetFocus()
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zkpc03(session, **kwargs) -> None:
    """
    Function to run zkpc03 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZKPC03"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/btn%_S_MATERI_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    try:
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        if 'file_path' in kwargs:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
        else:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
        if 'file_name' in kwargs:
            sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
        else:
            sapgui.sap_del_tmp_file()
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
        session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)
    except Exception as e:
        print(e)
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(3)


def cewb(session, **kwargs) -> None:
    """
    Function to run cewb code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: valid_date: str in format %y-%m-%d, id_list: dataframe, layout: str,
                             file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "CEWB"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/subWORKAREA:SAPLCPSC:1150/btnBUTTON_CHANGE").press()
    session.findById("wnd[1]").sendVKey(4)
    session.findById("wnd[2]").sendVKey(2)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'valid_date' in kwargs:
        session.findById("wnd[0]/usr/subVALIDITY:SAPLCPSC:2100/ctxtCWB_VALIDITY-DATUV").text = kwargs['valid_date']
        session.findById("wnd[0]/usr/subVALIDITY:SAPLCPSC:2100/ctxtCWB_VALIDITY-DATUV").SetFocus()
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/subSELECTION_CRITERIA:SAPLCPSC:1250/tabsTAB_STRIP_SEL/tabpBOMS/ssubSUBPAGE:SAPLCPSC:3340/btn%_MBMMATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[0]/btn[86]").press()
    if 'layout' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/lbl[1,1]").setFocus()
        session.findById("wnd[1]").sendVKey(71)
        session.findById("wnd[2]/usr/txtRSYSF-STRING").text = kwargs['layout']
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[3]/usr/lbl[1,2]").setFocus()
        session.findById("wnd[3]").sendVKey(2)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]").sendVKey(3)


def s_p99_41000062(session, **kwargs) -> None:
    """
    Function to run S_P99_41000062 code
    :param session:  parameter obtained from sapgui
    :param kwargs: optional: plant: str, id_list: dataframe, currency: str, layout: str,
                             file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "S_P99_41000062"
    session.findById("wnd[0]").sendVKey(0)
    if 'plant' in kwargs:
        session.findById("wnd[0]/usr/ctxtP_WERKS").text = kwargs['plant']
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/btn%_R_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'currency' in kwargs:
        session.findById("wnd[0]/usr/ctxtP_CURTP").text = kwargs['currency']
    if 'layout' in kwargs:
        session.findById("wnd[0]/usr/ctxtP_VARIAN").text = kwargs['layout']
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zc05h(session, **kwargs) -> None:
    """
    Function to run zc05H code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, variant_user: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zc05H"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[40]").press()
    if 'variant' in kwargs:
        session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = kwargs['variant']
    if 'variant_user' in kwargs:
        session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").text = kwargs['variant_user']
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zdisplay(session, **kwargs) -> None:
    """
    Function to run zdisplay code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: id_list: dataframe, file_path: str, file_name: str, cell_row: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zdisplay"
    session.findById("wnd[0]").sendVKey(0)
    if 'cell_row' in kwargs:
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = kwargs['cell_row']
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = kwargs['cell_row']
    else:
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = 6
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "6"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").clickCurrentCell()
    if 'id_list' in kwargs:
        print('test')
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/txtMAX_SEL").text = "99999999"
    session.findById("wnd[0]/usr/txtMAX_SEL").SetFocus()
    session.findById("wnd[0]").sendVKey(8)
    if 'cell_row' not in kwargs:
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = -1
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("MATNR")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&SORT_ASC")
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def mb5m(session, **kwargs) -> None:
    """
    Function to run mb5m code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "MB5M"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zcpo_mkal(session, **kwargs) -> None:
    """
    Function to run zcpo_mkal code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, valid_date: str in format %y-%m-%d, id_list: dataframe,
                             file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZCPO_MKAL"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'valid_date' in kwargs:
        session.findById("wnd[0]/usr/ctxtP_DATUV").text = kwargs['valid_date']
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").SetFocus()
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").SetFocus()
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def ppt1(session, **kwargs) -> None:
    """
    Function to run ppt1 code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, id_list: dataframe, layout: str,
                             date_from: str in format %y-%m-%d, date_to: str in format %y-%m-%d
                             file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n/sapapo/ppt1"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").Text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtSV_DTSTA").Text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtSV_DTEND").Text = kwargs['date_to']
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSELSCR1/ssub%_SUBSCREEN_SELBLOCK:/SAPAPO/SAPLRRP_PT_ENTRY:2010/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    if 'layout' in kwargs:
        session.findById("wnd[0]").sendVKey(kwargs['layout'])
    session.findById("wnd[0]/usr/subINCL1:/SAPAPO/SAPLRRP_FRAMES:7002/cntlCONTAINER/shellcont/shell/shellcont[0]/shell").pressToolbarButton("PTP_TOOLBAR_EXPAND")
    session.findById("wnd[0]/usr/subINCL1:/SAPAPO/SAPLRRP_FRAMES:7002/cntlCONTAINER/shellcont/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/subINCL1:/SAPAPO/SAPLRRP_FRAMES:7002/cntlCONTAINER/shellcont/shell/shellcont[0]/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    session.findById("wnd[0]").sendVKey(3)


def mb5s(session, **kwargs) -> None:
    """
    Function to run mb5s code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: id_list: dataframe, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "mb5s"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtEKORG-LOW").text = "PL01"
    session.findById("wnd[0]/usr/ctxtEKORG-LOW").SetFocus()
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/btn%_EBELN_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[43]").press()
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    session.findById("wnd[0]").sendVKey(3)


def vl06f(session, **kwargs) -> None:
    """
    Function to run vl06f code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, layout: str,
                             date_from: str in format %y-%m-%d, date_to: str in format %y-%m-%d
                             file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "VL06F"
    session.findById("wnd[0]").sendVKey(0)
    if 'variant' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtIT_ERDAT-LOW").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtIT_ERDAT-HIGH").text = kwargs['date_to']
    session.findById("wnd[0]/usr/ctxtIT_ERDAT-HIGH").SetFocus()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[18]").press()
    if 'layout' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").Selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]/tbar[0]/btn[12]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    session.findById("wnd[0]").sendVKey(3)


def zcpot_rrp(session, **kwargs) -> None:
    """
    Function to run zcpot_rrp_report code
    :param session: parameter obtained from sapgui
    :param kwargs: optional: variant: str, layout: str, id_list: dataframe
                             date_from: str in format %y-%m-%d,
                             date_to: str in format %y-%m-%d
                             file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = " ZCPOT_RRP_REPORT"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    if 'variant' in kwargs:
        session.findById("wnd[1]/usr/txtV-LOW").text = kwargs['variant']
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'id_list' in kwargs:
        kwargs['id_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSELSCR1/ssub%_SUBSCREEN_SELBLOCK:ZCPOR_RRP_REPORT:2010/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtP_DTSTA").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtP_DTEND").text = kwargs['date_to']
    session.findById("wnd[0]").sendVKey(8)
    if 'layout' in kwargs:
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = ""
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = kwargs['layout']
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").setFocus()
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]/tbar[0]/btn[12]").press()
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapguiparam.SAP_TMP_PATH
    if 'file_name' in kwargs:
        sapgui.sap_del_tmp_file(file_name=kwargs['file_name'])
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        sapgui.sap_del_tmp_file()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapguiparam.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)