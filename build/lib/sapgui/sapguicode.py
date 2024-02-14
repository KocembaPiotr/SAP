from sapgui import sapgui


def zmpua25(session, *args, **kwargs) -> None:
    """
    Function to run zmpua25 code
    :param session: parameter obtained from sapgui
    :param args: variant:str, product_id: DataFrame
    :return: None.
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMPUA25"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    args[1].to_clipboard(index=False, header=None)
    session.findById(r"wnd[0]/usr/btn%_R_MATNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById(r"wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    sapgui.sap_download_tmp_file_del()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def rrp4(session, *args, **kwargs) -> None:
    """
    Function to run rrp4 code
    :param session: parameter obtained from sapgui
    :param args: variant:str, dateFrom:str, dateTo:str,
                 product_id: DataFrame, layout: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/sapapo/RRP4"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]").sendVKey(8)
    session.findById("wnd[0]/usr/ctxtSV_DTSTA").text = args[1]
    session.findById("wnd[0]/usr/ctxtSV_DTEND").text = args[2]
    args[3].to_clipboard(index=False, header=None)
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSELSCR1/ssub%_SUBSCREEN_SELBLOCK:/SAPAPO/SAPLRRP_PT_ENTRY:2010/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
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
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarButton("&MB_VARIANT")
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = ""
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
    session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = args[4]
    session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").setFocus()
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[2]/tbar[0]/btn[12]").press()
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    sapgui.sap_download_tmp_file_del()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    try:
        session.findById("wnd[1]/usr/btnBUTTON_2").press()
    except Exception as e:
        print(e)
    session.findById("wnd[0]").sendVKey(3)


def rrp1(session, *args, **kwargs) -> None:
    """
    Function to run rrp1 code
    :param session: parameter obtained from sapgui
    :param args: variant:str, dateFrom:str, dateTo:str,
                 product_id: DataFrame, layout: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/sapapo/RRP1"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]").sendVKey(8)
    session.findById("wnd[0]/usr/ctxtSV_DTSTA").text = args[1]
    session.findById("wnd[0]/usr/ctxtSV_DTEND").text = args[2]
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSELSCR1/ssub%_SUBSCREEN_SELBLOCK:/SAPAPO/SAPLRRP_PT_ENTRY:2010/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()
    args[3].to_clipboard(index=False, header=None)
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
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarButton("&MB_VARIANT")
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellRow = -1
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectColumn("VARIANT")
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = ""
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").contextMenu()
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectContextMenuItem("&FIND")
    session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = args[4]
    session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").setFocus()
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[2]/tbar[0]/btn[12]").press()
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem("&PC")
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    sapgui.sap_download_tmp_file_del()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    try:
        session.findById("wnd[1]/usr/btnBUTTON_2").press()
    except Exception as e:
        print(e)
    session.findById("wnd[0]").sendVKey(3)


def zpp_mat(session, *args, **kwargs) -> None:
    """
    Function to run zpp_mat code
    :param session: parameter obtained from sapgui
    :param args: variant:str, variable: DataFrame according to option
    :param kwargs: option:str, to change function behaviour
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zpp_mat"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]").sendVKey(8)
    if 'option' not in kwargs or kwargs['option'] == 'ID':
        args[1].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    sapgui.sap_download_tmp_file_del()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def mb52(session, *args, **kwargs) -> None:
    """
    Function to run mb52 code
    :param session: parameter obtained from sapgui
    :param args: variant:str, variable: DataFrame according to option
    :param kwargs: option:str, to change function behaviour
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "mb52"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if 'option' in kwargs and kwargs['option'] == 'ID':
        args[1].to_clipboard(index=False, header=None)
        session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
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
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def me2m(session, *args, **kwargs) -> None:
    """
    Function to run me2m code
    :param session: parameter obtained from sapgui
    :param args: variant:str
    :param kwargs: optional variables: date_from, date_to in format %Y-%m-%d
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "me2m"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]").sendVKey(8)
    if 'date_from' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_EINDT-LOW").text = kwargs['date_from']
    if 'date_to' in kwargs:
        session.findById("wnd[0]/usr/ctxtS_EINDT-HIGH").text = kwargs['date_to']
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def m_ld(session, *args, **kwargs) -> None:
    """
        Function to run m_ld code
        :param session: parameter obtained from sapgui
        :param args: dataframe with IDs, date_from, date_to in format %Y-%m-%d : str
        :param kwargs: optional variables: file_path, file_name : str
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
    args[0].to_clipboard(index=False, header=None)
    session.findById("wnd[1]/tbar[0]/btn[16]").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtDATUM-LOW").text = args[1]
    session.findById("wnd[0]/usr/ctxtDATUM-HIGH").text = args[2]
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
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def mb51(session, *args, **kwargs) -> None:
    """
    Function to run MB51 code
    :param session: parameter obtained from sapgui
    :param args: variant_name: str
    :param kwargs: optional variables: id_list: dataframe, date_from: str in format %y-%m-%d,
                   date_to: str in format %y-%m-%d, batch_list: dataframe,
                   file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
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
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zmmla11(session, *args, **kwargs) -> None:
    """
    Function to run zmmla11 code
    :param session: parameter obtained from sapgui
    :param args: variant_name: str, product_id: DataFrame
    :param kwargs: optional variables: file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zmmla11"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]").sendVKey(8)
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
    args[1].to_clipboard(index=False, header=None)
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
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zka02(session, *args, **kwargs) -> None:
    """
    Function to run zka02 code
    :param session: parameter obtained from sapgui
    :param args: plant_name: str, variant_name: str, product_id: DataFrame
    :param kwargs: optional variables: file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZKA02"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtP_WERKS").text = args[0]
    session.findById("wnd[0]/usr/btn%P200010_1000").press()
    session.findById("wnd[1]/tbar[0]/btn[17]").press()
    session.findById("wnd[2]/usr/txtV-LOW").text = args[1]
    session.findById("wnd[2]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[2]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/usr/btn%_W_MATNR_%_APP_%-VALU_PUSH").press()
    args[2].to_clipboard(index=False, header=None)
    session.findById("wnd[2]/tbar[0]/btn[16]").press()
    session.findById("wnd[2]/tbar[0]/btn[24]").press()
    session.findById("wnd[2]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/usr/cntlCKKK_0200_CUSTOM_CTRL/shellcont/shell").selectAll()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    if 'file_path' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = kwargs['file_path']
    else:
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def coois(session, *args, **kwargs) -> None:
    """
    Function to run COOIS code
    :param session: parameter obtained from sapgui
    :param args: variant_name: str, id_list: dataframe
    :param kwargs: optional variables: file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "COOIS"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]").sendVKey(8)
    if 'order_list' in kwargs:
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press()
        kwargs['order_list'].to_clipboard(index=False, header=None)
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
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
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def zcp04(session, *args, **kwargs) -> None:
    """
    Function to run ZCP04 code
    :param session: parameter obtained from sapgui
    :param args: variant_name: str
    :param kwargs: optional variables: id_list: dataframe, date_from: str in format %y-%m-%d,
                   date_to: str in format %y-%m-%d, plant: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zcp04"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
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
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


def me5a(session, *args, **kwargs) -> None:
    """
    Function to run me5a code
    :param session: parameter obtained from sapgui
    :param args: variant: str, ids: DataFrame
    :param kwargs: optional variables: layout: str, file_path: str, file_name: str
    :return: None
    """
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "me5a"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = args[0]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/btn%_BA_BANFN_%_APP_%-VALU_PUSH").press()
    args[1].to_clipboard(index=False, header=None)
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
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sapgui.SAP_TMP_PATH
    if 'file_name' in kwargs:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = kwargs['file_name']
    else:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = sapgui.SAP_TMP_FILE
    session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "0000"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)