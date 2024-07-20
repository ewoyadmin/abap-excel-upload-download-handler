*&---------------------------------------------------------------------*
*& Report zzexcel_test
*&---------------------------------------------------------------------*
*& This report is an example of class zcl_excel_handler usage.
*&---------------------------------------------------------------------*
REPORT zzexcel_test.

PARAMETERS:
  p_ifile TYPE string LOWER CASE DEFAULT 'test.csv',
  p_ofile TYPE string LOWER CASE DEFAULT 'test_out',
  p_csv   RADIOBUTTON GROUP r1 DEFAULT 'X',
  p_xls   RADIOBUTTON GROUP r1.

TYPES:
  BEGIN OF ts_test,
    field1 TYPE string,
    field2 TYPE d,
    field3 TYPE p DECIMALS 2,
  END OF ts_test,
  tt_test TYPE STANDARD TABLE OF ts_test.

DATA: lt_test TYPE tt_test.

TRY.
    DATA(lo_excel) = NEW zcl_excel_handler( ).
    IF p_ifile CS '.xls'.
      lo_excel->upload_xlsx( EXPORTING iv_file_path = p_ifile ir_table = REF #( lt_test ) ).
    ELSE.
      lo_excel->upload_csv( EXPORTING iv_file_path = p_ifile ir_table = REF #( lt_test ) iv_hdr_lines = 1 ).
    ENDIF.
    IF p_ofile IS NOT INITIAL.
      IF p_csv EQ abap_true..
        DATA(lv_bytes) = lo_excel->download_csv( EXPORTING iv_file_path = |{ p_ofile }.csv| ir_table = REF #( lt_test ) ).
      ELSE.
        lv_bytes = lo_excel->download_xlsx( EXPORTING iv_file_path = |{ p_ofile }.xlsx| ir_table = REF #( lt_test ) ).
      ENDIF.
      MESSAGE s001(00) WITH lv_bytes ' bytes written to file ' p_ofile.
    ENDIF.

  CATCH zcx_excel_handler INTO DATA(lx_err).
    MESSAGE lx_err->get_text( ) TYPE 'I' DISPLAY LIKE 'E'.
ENDTRY.

cl_demo_output=>display( lt_test ).