*&---------------------------------------------------------------------*
*& Report ztest_excel_download
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT ztest_excel_download.

CLASS lcl_main DEFINITION.

  PUBLIC SECTION.

    CLASS-METHODS:
      get_file RETURNING VALUE(rv_file) TYPE string,
      export_file IMPORTING iv_data TYPE xstring
                            iv_file TYPE string OPTIONAL.

ENDCLASS.

CLASS lcl_main IMPLEMENTATION.

  METHOD get_file.

    DATA(title) = |Select Excel File, e.g. *.xlsx|.
    DATA(defaultextension) = |.xlsx|.
    DATA(filefilter) = `Excel Files (*.xlsx)|*.xlsx`.
    DATA it_tab TYPE filetable.
    DATA returncode TYPE i.

    CALL METHOD cl_gui_frontend_services=>file_open_dialog
      EXPORTING
        window_title      = title
        default_extension = defaultextension
      CHANGING
        file_table        = it_tab
        rc                = returncode.
    IF sy-subrc NE 0.
*   Implement suitable error handling here
    ENDIF.

    rv_file = VALUE #( it_tab[ 1 ] OPTIONAL ).

  ENDMETHOD.

  METHOD export_file.

    IF iv_file IS INITIAL.
      DATA(lv_pc) = abap_true.
      DATA(lv_filename) = lcl_main=>get_file(  ).
    ELSE.
      lv_pc = abap_false.
      lv_filename = iv_file.
    ENDIF.
    CHECK lv_filename IS NOT INITIAL.

    IF lv_pc EQ abap_true.
* Export to PC
      cl_scp_change_db=>xstr_to_xtab( EXPORTING im_xstring = iv_data
                                      IMPORTING ex_xtab    = DATA(filecontenttab) ).

      cl_gui_frontend_services=>gui_download(
        EXPORTING
          bin_filesize              = xstrlen( iv_data )
          filename                  = lv_filename
          filetype                  = 'BIN'
          confirm_overwrite         = abap_true
        IMPORTING
          filelength                = DATA(bytestransferred)
        CHANGING
          data_tab                  = filecontenttab
        EXCEPTIONS
          file_write_error          = 1
          no_batch                  = 2
          gui_refuse_filetransfer   = 3
          invalid_type              = 4
          no_authority              = 5
          unknown_error             = 6
          header_not_allowed        = 7
          separator_not_allowed     = 8
          filesize_not_allowed      = 9
          header_too_long           = 10
          dp_error_create           = 11
          dp_error_send             = 12
          dp_error_write            = 13
          unknown_dp_error          = 14
          access_denied             = 15
          dp_out_of_memory          = 16
          disk_full                 = 17
          dp_timeout                = 18
          file_not_found            = 19
          dataprovider_exception    = 20
          control_flush_error       = 21
          not_supported_by_gui      = 22
          error_no_gui              = 23
          OTHERS                    = 24
      ).
      IF sy-subrc NE 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                   WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ELSE.
        MESSAGE s001(00) WITH bytestransferred ' bytes transferred'.
      ENDIF.

    ELSE.
* Export to server
      OPEN DATASET lv_filename FOR OUTPUT IN BINARY MODE.
      IF sy-subrc EQ 0.
        TRANSFER iv_data TO lv_filename.
        CLOSE DATASET lv_filename.
        MESSAGE |File { lv_filename } downloaded to server successfully| TYPE 'S'.
      ENDIF.
    ENDIF.

  ENDMETHOD.

ENDCLASS.

PARAMETERS: p_file TYPE string LOWER CASE.

DATA:
  lt_flight TYPE TABLE OF /DMO/I_Flight_R.

START-OF-SELECTION.

  SELECT * FROM /DMO/I_Flight_R INTO TABLE @lt_flight.
  IF sy-subrc EQ 0.
    GET REFERENCE OF lt_flight INTO DATA(lo_data_ref).
    DATA(lv_xstring) = NEW zcl_itab_to_excel( )->itab_to_xstring( lo_data_ref ).

    lcl_main=>export_file( iv_data = lv_xstring
                           iv_file = p_file
                         ).

  ENDIF.





CLASS zcl_itab_to_excel DEFINITION PUBLIC FINAL.
  PUBLIC SECTION.
    METHODS:
      itab_to_xstring
        IMPORTING ir_data_ref       TYPE REF TO data
        RETURNING VALUE(rv_xstring) TYPE xstring.
ENDCLASS.

CLASS zcl_itab_to_excel IMPLEMENTATION.

  METHOD itab_to_xstring.

    FIELD-SYMBOLS: <fs_data> TYPE ANY TABLE.

    CLEAR rv_xstring.
    ASSIGN ir_data_ref->* TO <fs_data>.

    TRY.
        cl_salv_table=>factory(
          IMPORTING r_salv_table = DATA(lo_table)
          CHANGING  t_table      = <fs_data> ).

        DATA(lt_fcat) =
          cl_salv_controller_metadata=>get_lvc_fieldcatalog(
            r_columns      = lo_table->get_columns( )
            r_aggregations = lo_table->get_aggregations( ) ).

        DATA(lo_result) =
          cl_salv_ex_util=>factory_result_data_table(
            r_data         = ir_data_ref
            t_fieldcatalog = lt_fcat ).

        cl_salv_bs_tt_util=>if_salv_bs_tt_util~transform(
          EXPORTING
            xml_type      = if_salv_bs_xml=>c_type_xlsx
            xml_version   = cl_salv_bs_a_xml_base=>get_version( )
            r_result_data = lo_result
            xml_flavour   = if_salv_bs_c_tt=>c_tt_xml_flavour_export
            gui_type      = if_salv_bs_xml=>c_gui_type_gui
          IMPORTING
            xml           = rv_xstring ).
      CATCH cx_root.
        CLEAR rv_xstring.
    ENDTRY.
  ENDMETHOD.

ENDCLASS.