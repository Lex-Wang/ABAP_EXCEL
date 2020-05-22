FUNCTION ZBC_EXCEL_2_INNER_TABLE.
*"----------------------------------------------------------------------
*"*"Open Excel in OLE
*"  IMPORTING
*"     REFERENCE(PI_FILENAME) TYPE  STRING
*"     VALUE(PI_SHEETNAME) TYPE  STRING DEFAULT 'Sheet1'
*"  TABLES
*"      PT_TAB
*"  EXCEPTIONS
*"      FILE_OPEN_ERROR
*"----------------------------------------------------------------------
  types:
    lty_c30k(30000) type c.  

  TYPE-POOLS:
    ole2.

  DATA:
    ole_excel      TYPE ole2_object,
    ole_workbooks  TYPE ole2_object,
    ole_workbook   TYPE ole2_object,
    ole_worksheets TYPE ole2_object,
    ole_worksheet  TYPE ole2_object,
    ole_cell_begin TYPE ole2_object,
    ole_cell_end   TYPE ole2_object,
    ole_range      TYPE ole2_object.

  DATA:
    lv_subrc     TYPE sy-subrc,
    lv_begin_col TYPE i,
    lv_end_col   TYPE i,
    lv_begin_row TYPE i,
    lv_end_row   TYPE i,
    lv_add_rows  TYPE i VALUE 3000.

  DATA:
    lt_excel_tab     TYPE STANDARD TABLE OF lty_c30k,
    lw_excel_tab     TYPE lty_c30k,
    lw_excel_tab_tmp TYPE lty_c30k.


*->����Excel object
  CREATE OBJECT ole_excel 'Excel.Application'.
  IF sy-subrc <> 0.
    MESSAGE e001(ZBC) RAISING file_open_error.
  ENDIF.

*->
  GET PROPERTY OF ole_excel 'Workbooks' = ole_workbooks.

  CALL METHOD OF
      ole_workbooks
      'Open'        = ole_workbook
    EXPORTING
      #1            = pi_filename.

*->ȡ��Sheet
  if pi_sheetname is not INITIAL.
  GET PROPERTY OF ole_workbook 'Worksheets' = ole_worksheets
    EXPORTING
      #1 = pi_sheetname.

  IF sy-subrc = 0.
    CALL METHOD OF
      ole_worksheets
      'Activate'.
  ELSE.
    MESSAGE e003(zbc) RAISING file_open_error.
  ENDIF.
else.
  get PROPERTY of ole_workbook 'ActiveSheet' = ole_worksheets.
  if sy-subrc <> 0.
  message e003(zbc).
  endif.
endif.

*->��Sheet������Copy �� ClipBoard
  lv_begin_col = 1.
  lv_end_col = 256.
  lv_begin_row = 0.
  lv_end_row = 0.

  WHILE lv_subrc IS INITIAL.

    IF lv_begin_row IS INITIAL.
      lv_begin_row = 1.
      lv_end_row   = lv_add_rows.
    ELSE.
      lv_begin_row = lv_begin_row + lv_add_rows.
      lv_end_row   = lv_end_row   + lv_add_rows.
    ENDIF.

    CALL METHOD OF
        ole_worksheets
        'Cells'        = ole_cell_begin
      EXPORTING
        #1             = lv_begin_row
        #2             = lv_begin_col.

    CALL METHOD OF
        ole_worksheets
        'Cells'        = ole_cell_end
      EXPORTING
        #1             = lv_end_row
        #2             = lv_end_col.

    CALL METHOD OF
        ole_worksheets
        'RANGE'        = ole_range
      EXPORTING
        #1             = ole_cell_begin
        #2             = ole_cell_end.

    CALL METHOD OF
      ole_range
      'SELECT'.
    IF sy-subrc <> 0.
      EXIT.
    ENDIF.

    CALL METHOD OF
      ole_range
      'COPY'.

* read clipboard into ABAP
    CALL METHOD cl_gui_frontend_services=>clipboard_import
      IMPORTING
        data                 = lt_excel_tab
      EXCEPTIONS
        cntl_error           = 1
        error_no_gui         = 2
        not_supported_by_gui = 3
        OTHERS               = 4.
    IF sy-subrc <> 0.
      MESSAGE 'Error during import of clipboard contents' TYPE 'A'.
    ENDIF.

    lv_subrc = 4.
    LOOP AT lt_excel_tab INTO lw_excel_tab.
      lw_excel_tab_tmp = lw_excel_tab.
      REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>horizontal_tab IN lw_excel_tab_tmp WITH space.
      IF NOT ( lw_excel_tab_tmp = space OR lw_excel_tab_tmp IS INITIAL ).
        APPEND lw_excel_tab TO pt_tab.
        CLEAR lv_subrc.
      ENDIF.
    ENDLOOP.

    CLEAR lt_excel_tab.

  ENDWHILE.

  call METHOD of ole_workbook
    'Close'.

  CALL METHOD OF
    ole_excel
    'QUIT'.
  FREE:
    ole_excel      ,
    ole_workbooks  ,
    ole_workbook   ,
    ole_worksheets ,
    ole_worksheet  ,
    ole_cell_begin ,
    ole_cell_end   ,
    ole_range      .


ENDFUNCTION.