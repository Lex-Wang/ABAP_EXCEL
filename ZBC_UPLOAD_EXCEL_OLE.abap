FUNCTION ZBC_UPLOAD_EXCEL_OLE.
*"----------------------------------------------------------------------
*"*"���ؽӿڣ�
*"  IMPORTING
*"     VALUE(PI_FILENAME) TYPE  STRING
*"     VALUE(PI_SHEETNAME) TYPE  STRING DEFAULT 'Sheet1'
*"     VALUE(PI_STARTLINE) TYPE  I DEFAULT '1'
*"     VALUE(PI_STARTCOLUMN) TYPE  I DEFAULT '1'
*"     VALUE(PI_SKPCL_TBL) TYPE  I DEFAULT '0'
*"  TABLES
*"      PT_TAB
*"  EXCEPTIONS
*"      OPEN_FILE_ERR
*"----------------------------------------------------------------------
  types:
    lty_c30K(30000) type c.

  DATA:
    lt_tabc       TYPE STANDARD TABLE OF lty_c30K,
    lw_tabc       TYPE lty_c30K,
    lv_tabix      TYPE sy-tabix,
    lt_cell       TYPE STANDARD TABLE OF string,
    lw_cell       TYPE string,
    lv_cell_tabix TYPE sy-tabix.

  DATA:
    lv_column_num TYPE i,
    lv_column_skp TYPE i,
    lw_tab_ref    TYPE REF TO data.

  data:
    lo_cx_root type REF TO cx_root.

  FIELD-SYMBOLS:
    <lw_tab>   TYPE any,
    <lv_value> TYPE any.

  CREATE DATA lw_tab_ref LIKE LINE OF pt_tab.
  ASSIGN lw_tab_ref->* TO <lw_tab>.

  lv_column_skp = pi_skpcl_tbl.

*->�����ݷŵ��ڱ�
  CALL FUNCTION 'ZBC_EXCEL_2_INNER_TABLE'
    EXPORTING
      pi_filename     = pi_filename
      pi_sheetname    = pi_sheetname
    TABLES
      pt_tab          = lt_tabc
    EXCEPTIONS
      file_open_error = 1
      OTHERS          = 2.
  IF sy-subrc <> 0.
    MESSAGE e002(zbc) RAISING open_file_err.
  ENDIF.

*->�������и�ڱ�
  LOOP AT lt_tabc INTO lw_tabc.
    lv_tabix = sy-tabix.

    CHECK lv_tabix >= pi_startline.
    CLEAR lt_cell.

    SPLIT lw_tabc AT cl_abap_char_utilities=>horizontal_tab INTO TABLE lt_cell.

    LOOP AT lt_cell INTO lw_cell.
      lv_cell_tabix = sy-tabix.

      lv_column_num = lv_cell_tabix - pi_startcolumn + 1 + lv_column_skp.

      CHECK lv_column_num > 0.

      ASSIGN COMPONENT lv_column_num OF STRUCTURE <lw_tab> TO <lv_value>.
      CHECK sy-subrc = 0.
      TRY.
          <lv_value> = lw_cell.
        CATCH cx_root INTO lo_cx_root.
          MESSAGE e001(zbc) with lv_tabix lv_column_num RAISING open_file_err.
      ENDTRY.
    ENDLOOP.
    IF <lw_tab> IS NOT INITIAL.
      APPEND <lw_tab> TO pt_tab.
      CLEAR <lw_tab>.
    ENDIF.

  ENDLOOP.
ENDFUNCTION.