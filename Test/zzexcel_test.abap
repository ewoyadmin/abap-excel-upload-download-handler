*&---------------------------------------------------------------------*
*& Report zzexcel_test
*&---------------------------------------------------------------------*
*& This report is an example of class zcl_excel_handler usage.
*&---------------------------------------------------------------------*
REPORT zzexcel_test.

PARAMETERS:
  p_ifile  TYPE string LOWER CASE DEFAULT 'test.xlsx',
  p_ofile  TYPE string LOWER CASE DEFAULT 'test_out',
  p_xls    RADIOBUTTON GROUP r1 DEFAULT 'X',
  p_csv    RADIOBUTTON GROUP r1,
  p_server AS CHECKBOX,
  p_conv   AS CHECKBOX.

TYPES:
  BEGIN OF ts_test,
    field1 TYPE string,
    field2 TYPE sy-datum,
    field3 TYPE p DECIMALS 2,
    field4 TYPE cc_sernr,
    field5 TYPE string,
  END OF ts_test,
  tt_test TYPE STANDARD TABLE OF ts_test.

DATA:
  lt_test  TYPE tt_test,
  lv_index TYPE string.

TRY.
    DATA(lo_excel) = NEW zcl_excel_handler( ).
    IF p_ifile IS NOT INITIAL.
      IF p_ifile CS '.xls'.
        lo_excel->upload_xlsx( EXPORTING iv_file_path = p_ifile ir_table = REF #( lt_test ) iv_hdr_lines = 0 iv_server = p_server ).
      ELSE.
        lo_excel->upload_csv( EXPORTING iv_file_path = p_ifile ir_table = REF #( lt_test ) iv_hdr_lines = 0 iv_server = p_server ).
      ENDIF.
    ENDIF.
    IF p_ofile IS NOT INITIAL.
      IF lt_test IS INITIAL.
        DO 1000 TIMES.
          lv_index = sy-index.
          lt_test[] = VALUE #( BASE lt_test[]
                               ( field1 = 'test' field2 = sy-datum field3 = '123456.78' field4 = '12345' field5 = 'tuitui' ) ).
        ENDDO.
      ENDIF.
      IF p_csv EQ abap_true.
        DATA(lv_bytes) = lo_excel->download_csv( EXPORTING iv_file_path = |{ p_ofile }.csv| ir_table = REF #( lt_test ) iv_server = p_server ).
      ELSE.
        lv_bytes = lo_excel->download_xlsx( EXPORTING iv_file_path = |{ p_ofile }.xlsx| ir_table = REF #( lt_test ) iv_server = p_server ).
      ENDIF.
      MESSAGE s001(00) WITH lv_bytes ' bytes written to file ' p_ofile.
    ENDIF.

  CATCH zcx_excel_handler INTO DATA(lx_err).
    MESSAGE lx_err->get_text( ) TYPE 'I' DISPLAY LIKE 'E'.
ENDTRY.

IF lt_test IS NOT INITIAL.
  IF sy-batch EQ abap_false.
    cl_demo_output=>display( lt_test ).
  ENDIF.
  LOOP AT lt_test INTO DATA(ls_test).
    DO.
      ASSIGN COMPONENT sy-index OF STRUCTURE ls_test TO FIELD-SYMBOL(<fs_fld>).
      IF sy-subrc NE 0.
        EXIT.
      ENDIF.
      IF p_conv EQ abap_true.
        WRITE: <fs_fld>.
      ELSE.
        WRITE: <fs_fld> USING NO EDIT MASK.
      ENDIF.
    ENDDO.
    NEW-LINE.
  ENDLOOP.
ENDIF.