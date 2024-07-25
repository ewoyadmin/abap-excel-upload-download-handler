"! General Excel upload/download handler class
"! <p>
"! This class supports XLSX and CSV file formats for uploading and downloading data
"! from/to SAP from/to PC and application server.
"! </p>
CLASS zcl_excel_handler DEFINITION PUBLIC FINAL CREATE PUBLIC.

  PUBLIC SECTION.
    "! Constructor
    "! @parameter iv_separator | Column separator for CSV files (default is semicolon)
    "! @raising zcx_excel_handler | Exception if invalid separator is provided
    METHODS constructor
      IMPORTING
        VALUE(iv_separator) TYPE c DEFAULT ';'
      RAISING
        zcx_excel_handler.

    "! Upload CSV file
    "! @parameter iv_file_path | Path to the CSV file
    "! @parameter iv_server | Flag to indicate if file is on application server
    "! @parameter iv_hdr_lines | Number of header lines to skip
    "! @parameter ir_table | Reference to the internal table to store data
    "! @raising zcx_excel_handler | Exception if upload fails
    METHODS upload_csv
      IMPORTING
        VALUE(iv_file_path) TYPE string OPTIONAL
        VALUE(iv_server)    TYPE abap_bool DEFAULT abap_false
        VALUE(iv_hdr_lines) TYPE i DEFAULT 1
        ir_table            TYPE REF TO data
      RAISING
        zcx_excel_handler.

    "! Download CSV file
    "! @parameter iv_file_path | Path to save the CSV file
    "! @parameter iv_server | Flag to indicate if file should be saved on application server
    "! @parameter iv_header | Flag to include header in CSV
    "! @parameter ir_table | Reference to the internal table with data
    "! @parameter rv_bytes_written | Number of bytes written
    "! @raising zcx_excel_handler | Exception if download fails
    METHODS download_csv
      IMPORTING
        VALUE(iv_file_path)     TYPE string OPTIONAL
        VALUE(iv_server)        TYPE abap_bool DEFAULT abap_false
        VALUE(iv_header)        TYPE abap_bool DEFAULT abap_true
        ir_table                TYPE REF TO data
      RETURNING
        VALUE(rv_bytes_written) TYPE i
      RAISING
        zcx_excel_handler.

    "! Upload XLSX file
    "! @parameter iv_file_path | Path to the XLSX file
    "! @parameter iv_server | Flag to indicate if file is on application server
    "! @parameter ir_table | Reference to the internal table to store data
    "! @raising zcx_excel_handler | Exception if upload fails
    METHODS upload_xlsx
      IMPORTING
        VALUE(iv_file_path) TYPE string OPTIONAL
        VALUE(iv_server)    TYPE abap_bool DEFAULT abap_false
        ir_table            TYPE REF TO data
      RAISING
        zcx_excel_handler.

    "! Download XLSX file
    "! @parameter iv_file_path | Path to save the XLSX file
    "! @parameter iv_server | Flag to indicate if file should be saved on application server
    "! @parameter ir_table | Reference to the internal table with data
    "! @parameter rv_bytes_written | Number of bytes written
    "! @raising zcx_excel_handler | Exception if download fails
    METHODS download_xlsx
      IMPORTING
        VALUE(iv_file_path)     TYPE string OPTIONAL
        VALUE(iv_server)        TYPE abap_bool DEFAULT abap_false
        ir_table                TYPE REF TO data
      RETURNING
        VALUE(rv_bytes_written) TYPE i
      RAISING
        zcx_excel_handler.

  PRIVATE SECTION.

    CONSTANTS:
      BEGIN OF mc,
        BEGIN OF msgty,
          success TYPE msgty VALUE 'S',
          error   TYPE msgty VALUE 'E',
          warning TYPE msgty VALUE 'W',
          info    TYPE msgty VALUE 'I',
          abend   TYPE msgty VALUE 'A',
        END OF msgty,
        is_numeric TYPE char11 VALUE '1234567890 ',
      END OF mc.

    TYPES: tt_text_data TYPE STANDARD TABLE OF text4096.

    DATA:
      mv_separator    TYPE c,
      mv_use_excel    TYPE abap_bool,
      mo_table_descr  TYPE REF TO cl_abap_tabledescr,
      mo_struct_descr TYPE REF TO cl_abap_structdescr.

    "! Get file path through file dialog
    "! @parameter iv_xlsx | Flag to indicate if file is XLSX
    "! @parameter rv_file | Selected file path
    METHODS get_file
      IMPORTING
        iv_xlsx        TYPE abap_bool DEFAULT abap_false
      RETURNING
        VALUE(rv_file) TYPE string.

    "! Get table structure
    "! @parameter ir_table | Reference to the internal table
    "! @parameter rt_components | Table of structure components
    METHODS get_table_structure
      IMPORTING
        ir_table             TYPE ANY TABLE
      RETURNING
        VALUE(rt_components) TYPE cl_abap_structdescr=>component_table.

    "! Convert line to structure
    "! @parameter iv_line | Input line
    "! @parameter it_components | Table of structure components
    "! @parameter ir_line | Reference to the structure
    "! @raising zcx_excel_handler | Exception if conversion fails
    METHODS convert_line_to_structure
      IMPORTING
        VALUE(iv_line)       TYPE string
        VALUE(it_components) TYPE cl_abap_structdescr=>component_table
        ir_line              TYPE REF TO data
      RAISING
        zcx_excel_handler.

    "! Convert structure to line
    "! @parameter ir_structure | Reference to the structure
    "! @parameter it_components | Table of structure components
    "! @parameter rv_line | Resulting line
    "! @raising zcx_excel_handler | Exception if conversion fails
    METHODS convert_structure_to_line
      IMPORTING
        ir_structure         TYPE any
        VALUE(it_components) TYPE cl_abap_structdescr=>component_table
      RETURNING
        VALUE(rv_line)       TYPE string
      RAISING
        zcx_excel_handler.

    "! Convert internal table to XLSX format
    "! @parameter ir_data_ref | Reference to the internal table
    "! @parameter rv_xstring | Resulting XLSX data as XSTRING
    METHODS itab_to_xlsx
      IMPORTING
        ir_data_ref       TYPE REF TO data
      RETURNING
        VALUE(rv_xstring) TYPE xstring.

    "! Validate number string
    "! @parameter cv_number_str | Number string to validate
    "! @parameter rv_is_valid | Flag indicating if number is valid
    METHODS validate_number
      CHANGING
        cv_number_str      TYPE string
      RETURNING
        VALUE(rv_is_valid) TYPE abap_bool.

    "! Check if running on Windows
    "! @parameter rv_result | True if running on Windows
    METHODS is_windows
      RETURNING
        VALUE(rv_result) TYPE abap_bool.

    "! Generate header line
    "! @parameter it_components | Table of structure components
    "! @parameter rv_line | Resulting header line
    METHODS header_line
      IMPORTING
        VALUE(it_components) TYPE cl_abap_structdescr=>component_table
      RETURNING
        VALUE(rv_line)       TYPE string.

ENDCLASS.


CLASS zcl_excel_handler IMPLEMENTATION.


  METHOD constructor.

    " Validate the column separator for CSV files
    CASE iv_separator.
      WHEN ','
        OR ';'
        OR cl_abap_char_utilities=>horizontal_tab.
        mv_separator = iv_separator.
      WHEN OTHERS.
        RAISE EXCEPTION TYPE zcx_excel_handler
          EXPORTING
            textid = zcx_excel_handler=>invalid_delimiter.
    ENDCASE.

  ENDMETHOD.


  METHOD convert_line_to_structure.

    DATA:
      lt_values TYPE TABLE OF string,
      lv_value  TYPE string,
      lv_data   TYPE REF TO data,
      lv_datfm  TYPE xudatfm.

    FIELD-SYMBOLS:
      <fs_line>  TYPE any,
      <fs_value> TYPE any.


    ASSIGN ir_line->* TO <fs_line>.

    SPLIT iv_line AT mv_separator INTO TABLE lt_values.

    " Raise exception if separator is found in column's value
    IF lines( it_components ) NE lines( lt_values ).
      RAISE EXCEPTION TYPE zcx_excel_handler
          EXPORTING
            textid = zcx_excel_handler=>invalid_delimiter.
    ENDIF.

    LOOP AT it_components INTO DATA(ls_component).
      lv_value = VALUE #( lt_values[ sy-tabix ] OPTIONAL ).
      ASSIGN COMPONENT ls_component-name OF STRUCTURE <fs_line> TO <fs_value>.
      IF sy-subrc EQ 0.
        DESCRIBE FIELD <fs_value> TYPE DATA(lv_type).
        CASE lv_type.
          WHEN 'P'.   "Packed number
            IF validate_number( CHANGING cv_number_str = lv_value ) EQ abap_false.
              CLEAR lv_value.
            ENDIF.

          WHEN 'D'.   "Date
            DO 6 TIMES.
              lv_datfm = sy-index.
              TRY.
                  cl_abap_datfm=>conv_date_ext_to_int(
                    EXPORTING
                      im_datext    = lv_value
                      im_datfmdes  = lv_datfm
                    IMPORTING
                      ex_datint    = DATA(lv_date)
                  ).
                CATCH cx_root.
                  CLEAR lv_date.
                  CONTINUE.   "Try all formats
              ENDTRY.
              lv_value = lv_date.
              EXIT.
            ENDDO.
            IF lv_date IS INITIAL.
              RAISE EXCEPTION TYPE zcx_excel_handler
                EXPORTING
                  textid = zcx_excel_handler=>invalid_date
                  msgv1  = CONV #( lv_value ).
            ENDIF.
        ENDCASE.
        TRY.
            <fs_value> = lv_value.
          CATCH cx_root.
            RAISE EXCEPTION TYPE zcx_excel_handler
              EXPORTING
                textid = zcx_excel_handler=>invalid_value
                msgv1  = CONV #( lv_value )
                msgv2  = CONV #( ls_component-name ).
        ENDTRY.
      ELSE.
        RAISE EXCEPTION TYPE zcx_excel_handler
          EXPORTING
            textid = zcx_excel_handler=>column_not_found
            msgv1  = CONV #( ls_component-name ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


  METHOD convert_structure_to_line.

    FIELD-SYMBOLS: <fs_component> TYPE any.

    LOOP AT it_components INTO DATA(ls_component).
      ASSIGN COMPONENT ls_component-name OF STRUCTURE ir_structure TO <fs_component>.
      IF sy-subrc EQ 0.
        IF rv_line IS INITIAL.
          rv_line = <fs_component>.
        ELSE.
          rv_line = rv_line && mv_separator && <fs_component>.
        ENDIF.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.

  METHOD header_line.

    LOOP AT it_components INTO DATA(ls_component).
      IF rv_line IS INITIAL.
        rv_line = ls_component-name.
      ELSE.
        rv_line = rv_line && mv_separator && ls_component-name.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.

  METHOD download_csv.

    DATA: lt_components TYPE cl_abap_structdescr=>component_table,
          lv_line       TYPE string,
          lr_table      TYPE REF TO data,
          lt_download   TYPE TABLE OF string,
          lv_errmsg     TYPE bapi_msg.

    FIELD-SYMBOLS:
      <fs_table> TYPE STANDARD TABLE,
      <fs_line>  TYPE any.


    ASSIGN ir_table->* TO <fs_table>.

    lt_components = get_table_structure( <fs_table> ).

    IF iv_server EQ abap_true.

      " Download to server disk
      OPEN DATASET iv_file_path FOR OUTPUT IN TEXT MODE ENCODING DEFAULT MESSAGE lv_errmsg.
      IF sy-subrc NE 0.
        RAISE EXCEPTION TYPE zcx_excel_handler
          EXPORTING
            textid = zcx_excel_handler=>file_write_error
            msgv1  = CONV #( iv_file_path ).
      ENDIF.

      DATA(lv_bytes) = 0.
      IF iv_header EQ abap_true.
        lv_line = header_line( lt_components ).
        TRANSFER lv_line TO iv_file_path LENGTH lv_bytes.
        rv_bytes_written = rv_bytes_written + lv_bytes.
      ENDIF.

      LOOP AT <fs_table> ASSIGNING <fs_line>.
        lv_line = convert_structure_to_line( ir_structure = <fs_line> it_components = lt_components ).
        TRANSFER lv_line TO iv_file_path LENGTH lv_bytes.
        rv_bytes_written = rv_bytes_written + lv_bytes.
      ENDLOOP.

      CLOSE DATASET iv_file_path.

    ELSE.

      " Prompt for file name when not supplier
      IF iv_file_path IS INITIAL.
        iv_file_path = get_file( ).
      ENDIF.

      IF iv_header EQ abap_true.
        lv_line = header_line( lt_components ).
        APPEND lv_line TO lt_download.
      ENDIF.

      " Convert table to CSV
      LOOP AT <fs_table> ASSIGNING <fs_line>.
        lv_line = convert_structure_to_line( ir_structure = <fs_line> it_components = lt_components ).
        APPEND lv_line TO lt_download.
      ENDLOOP.

      " Download to workstation
      cl_gui_frontend_services=>gui_download(
        EXPORTING
          filename                  = iv_file_path
          filetype                  = 'ASC'
          confirm_overwrite         = abap_true
        IMPORTING
          filelength                = DATA(lv_bytestransferred)
        CHANGING
          data_tab                  = lt_download
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
      IF sy-subrc EQ 0.
        rv_bytes_written = lv_bytestransferred.
      ELSE.
        RAISE EXCEPTION TYPE zcx_excel_handler
          EXPORTING
            textid = zcx_excel_handler=>file_write_error
            msgv1  = CONV #( iv_file_path ).
      ENDIF.

    ENDIF.

  ENDMETHOD.


  METHOD download_xlsx.

    DATA: lt_components TYPE cl_abap_structdescr=>component_table,
          lv_line       TYPE string,
          lr_table      TYPE REF TO data,
          lt_download   TYPE TABLE OF string,
          lv_errmsg     TYPE bapi_msg.

    FIELD-SYMBOLS:
      <fs_table> TYPE STANDARD TABLE,
      <fs_line>  TYPE any.


    ASSIGN ir_table->* TO <fs_table>.

    " Convert internal table into XLSX format
    DATA(lx_excel_data) = itab_to_xlsx( ir_table ).
    DATA(lv_binlen) = xstrlen( lx_excel_data ).

    IF iv_server EQ abap_true.

      " Download to server disk
      OPEN DATASET iv_file_path FOR OUTPUT IN BINARY MODE MESSAGE lv_errmsg.
      IF sy-subrc NE 0.
        RAISE EXCEPTION TYPE zcx_excel_handler
          EXPORTING
            textid = zcx_excel_handler=>file_write_error
            msgv1  = CONV #( iv_file_path ).
      ENDIF.

      TRANSFER lx_excel_data TO iv_file_path LENGTH lv_binlen.

      CLOSE DATASET iv_file_path.

      rv_bytes_written = rv_bytes_written + lv_binlen.

    ELSE.

      " Prompt for file name when not supplier
      IF iv_file_path IS INITIAL.
        iv_file_path = get_file( ).
      ENDIF.

      " Convert XSTRING to XSTRING table
      cl_scp_change_db=>xstr_to_xtab( EXPORTING im_xstring = lx_excel_data
                                      IMPORTING ex_xtab    = DATA(lv_filecontenttab) ).

      " Download as Excel file on workstation
      cl_gui_frontend_services=>gui_download(
        EXPORTING
          bin_filesize              = lv_binlen
          filename                  = iv_file_path
          filetype                  = 'BIN'
          confirm_overwrite         = abap_true
        IMPORTING
          filelength                = DATA(lv_bytestransferred)
        CHANGING
          data_tab                  = lv_filecontenttab
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
      IF sy-subrc EQ 0.
        rv_bytes_written = lv_bytestransferred.
      ELSE.
        RAISE EXCEPTION TYPE zcx_excel_handler
          EXPORTING
            textid = zcx_excel_handler=>file_write_error
            msgv1  = CONV #( iv_file_path ).
      ENDIF.
    ENDIF.

  ENDMETHOD.


  METHOD get_file.

    DATA:
      it_tab     TYPE filetable,
      returncode TYPE i.

    IF iv_xlsx EQ abap_true.
      DATA(title)            = |Select Excel File, e.g. *.xlsx|.
      DATA(defaultextension) = |.xlsx|.
      DATA(filefilter)       = `Excel Files (*.xlsx)|*.xlsx`.
    ELSE.
      title            = |Select CSV File, e.g. *.csv|.
      defaultextension = |.csv|.
      filefilter       = `Excel Files (*.csv)|*.csv`.
    ENDIF.

    CALL METHOD cl_gui_frontend_services=>file_open_dialog
      EXPORTING
        window_title      = title
        default_extension = defaultextension
      CHANGING
        file_table        = it_tab
        rc                = returncode.
    IF sy-subrc NE 0.
      " Implement suitable error handling here
    ENDIF.

    rv_file = VALUE #( it_tab[ 1 ] OPTIONAL ).

  ENDMETHOD.


  METHOD get_table_structure.

    mo_table_descr ?= cl_abap_structdescr=>describe_by_data( ir_table ).
    mo_struct_descr ?= mo_table_descr->get_table_line_type( ).
    rt_components = mo_struct_descr->get_components( ).

  ENDMETHOD.


  METHOD itab_to_xlsx.

    FIELD-SYMBOLS: <fs_data> TYPE ANY TABLE.

    CLEAR rv_xstring.
    ASSIGN ir_data_ref->* TO <fs_data>.

    " Convert internal table to XLSX format
    TRY.
        cl_salv_table=>factory(
          IMPORTING r_salv_table = DATA(lo_table)
          CHANGING  t_table      = <fs_data> ).

        " Get field catalog
        DATA(lt_fcat) =
          cl_salv_controller_metadata=>get_lvc_fieldcatalog(
            r_columns      = lo_table->get_columns( )
            r_aggregations = lo_table->get_aggregations( ) ).


        " If fields are not in DDIC, the add field names as header
        LOOP AT lt_fcat REFERENCE INTO DATA(lr_fcat)
             WHERE seltext   IS INITIAL
               AND scrtext_s IS INITIAL
               AND scrtext_m IS INITIAL
               AND scrtext_l IS INITIAL.
          lr_fcat->seltext = lr_fcat->fieldname.
        ENDLOOP.

        DATA(lo_result) =
          cl_salv_ex_util=>factory_result_data_table(
            r_data         = ir_data_ref
            t_fieldcatalog = lt_fcat ).

        " Create XLSX formatted data
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


  METHOD upload_csv.

    DATA:
      lv_line           TYPE string,
      lt_components     TYPE cl_abap_structdescr=>component_table,
      lv_num_of_records TYPE i,
      lo_struct_descr   TYPE REF TO cl_abap_structdescr,
      lr_table          TYPE REF TO data,
      lt_upload         TYPE STANDARD TABLE OF string,
      lv_errmsg         TYPE bapi_msg.

    FIELD-SYMBOLS:
      <fs_table> TYPE STANDARD TABLE,
      <fs_line>  TYPE any.


    ASSIGN ir_table->* TO <fs_table>.
    REFRESH <fs_table>.

    lt_components = get_table_structure( <fs_table> ).

    IF iv_server EQ abap_true.

      " Upload from application server disk
      OPEN DATASET iv_file_path FOR INPUT IN TEXT MODE ENCODING DEFAULT MESSAGE lv_errmsg.
      IF sy-subrc NE 0.
        RAISE EXCEPTION TYPE zcx_excel_handler
          EXPORTING
            textid = zcx_excel_handler=>file_not_found
            msgv1  = CONV #( iv_file_path )
            msgv2  = CONV #( lv_errmsg ).
      ENDIF.
      DATA(lv_row) = 0.
      DO.
        READ DATASET iv_file_path INTO lv_line.
        IF sy-subrc NE 0.
          EXIT.
        ENDIF.
        lv_row = lv_row + 1.
        CHECK lv_row GT iv_hdr_lines.
        APPEND INITIAL LINE TO <fs_table> ASSIGNING <fs_line>.
        DATA(lv_ref) = REF #( <fs_line> ).
        convert_line_to_structure( EXPORTING iv_line = lv_line it_components = lt_components ir_line = lv_ref ).
      ENDDO.

      CLOSE DATASET iv_file_path.

    ELSE.

      " Prompt for file name when not supplier
      IF iv_file_path IS INITIAL.
        iv_file_path = get_file( ).
      ENDIF.

      " Upload from workstation
      cl_gui_frontend_services=>gui_upload(
        EXPORTING
          filename                = iv_file_path
          filetype                = 'ASC'
        CHANGING
          data_tab                = lt_upload
        EXCEPTIONS
          file_open_error         = 1                " File does not exist and cannot be opened
          file_read_error         = 2                " Error when reading file
          no_batch                = 3                " Cannot execute front-end function in background
          gui_refuse_filetransfer = 4                " Incorrect front end or error on front end
          invalid_type            = 5                " Incorrect parameter FILETYPE
          no_authority            = 6                " No upload authorization
          unknown_error           = 7                " Unknown error
          bad_data_format         = 8                " Cannot Interpret Data in File
          header_not_allowed      = 9                " Invalid header
          separator_not_allowed   = 10               " Invalid separator
          header_too_long         = 11               " Header information currently restricted to 1023 bytes
          unknown_dp_error        = 12               " Error when calling data provider
          access_denied           = 13               " Access to File Denied
          dp_out_of_memory        = 14               " Not enough memory in data provider
          disk_full               = 15               " Storage medium is full.
          dp_timeout              = 16               " Data provider timeout
          not_supported_by_gui    = 17               " GUI does not support this
          error_no_gui            = 18               " GUI not available
          OTHERS                  = 19
      ).
      IF sy-subrc EQ 0.
        LOOP AT lt_upload INTO lv_line.
          CHECK sy-tabix GT iv_hdr_lines.
          APPEND INITIAL LINE TO <fs_table> ASSIGNING <fs_line>.
          lv_ref = REF #( <fs_line> ).
          convert_line_to_structure( EXPORTING iv_line = lv_line it_components = lt_components ir_line = lv_ref ).
        ENDLOOP.
      ELSE.
        RAISE EXCEPTION TYPE zcx_excel_handler
          EXPORTING
            textid = zcx_excel_handler=>file_not_found
            msgv1  = CONV #( iv_file_path ).
      ENDIF.

    ENDIF.

  ENDMETHOD.

  METHOD validate_number.

    DATA: lv_number TYPE p DECIMALS 2.

    rv_is_valid = abap_false.

    " Define a regex patterns for the number format with thousands separators and decimals.
    CONSTANTS:
      lc_pattern1 TYPE string VALUE '^\d{1,3}(,\d{3})*(\.\d+)?$',       "Comma as thousand's separator
      lc_pattern2 TYPE string VALUE '^\d{1,3}(.\d{3})*(\,\d+)?$'.       "Period as thousand's separator

    FIND PCRE lc_pattern1 IN cv_number_str.
    IF sy-subrc EQ 0.
      " Remove commas (thousands separators) to get a clean number string
      REPLACE ALL OCCURRENCES OF ',' IN cv_number_str WITH ''.
    ELSE.
      FIND PCRE lc_pattern2 IN cv_number_str.
      IF sy-subrc EQ 0.
        " Remove periods (thousands separators) to get a clean number string
        REPLACE ALL OCCURRENCES OF '.' IN cv_number_str WITH ''.
        TRANSLATE cv_number_str USING ',.'.     "Period as decimal point
      ENDIF.
    ENDIF.


    " Convert the cleaned string to a packed number
    TRY.
        CONDENSE cv_number_str NO-GAPS.
        lv_number = cv_number_str.
        rv_is_valid = abap_true.
      CATCH cx_sy_conversion_error.
        rv_is_valid = abap_false.
    ENDTRY.

  ENDMETHOD.


  METHOD upload_xlsx.

    DATA:
      lt_raw    TYPE tt_text_data,
      lv_errmsg TYPE bapi_msg.

    FIELD-SYMBOLS:
      <fs_table> TYPE STANDARD TABLE.


    IF NOT is_windows( ).
      RAISE EXCEPTION TYPE zcx_excel_handler
        EXPORTING
          textid = zcx_excel_handler=>not_supported.
    ENDIF.

    IF iv_server EQ abap_true.
      RETURN.   "Only from workstation
    ENDIF.

    " Prompt for file name when not supplier
    IF iv_file_path IS INITIAL.
      iv_file_path = get_file( ).
    ENDIF.

    ASSIGN ir_table->* TO <fs_table>.

    " Upload Excel format file (only on Windows)
    CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
      EXPORTING
        i_line_header        = abap_true
        i_tab_raw_data       = lt_raw
        i_filename           = CONV localfile( iv_file_path )
      TABLES
        i_tab_converted_data = <fs_table>
      EXCEPTIONS
        conversion_failed    = 1
        OTHERS               = 2.
    IF sy-subrc NE 0.
      RAISE EXCEPTION TYPE zcx_excel_handler
        EXPORTING
          textid = zcx_excel_handler=>conversion_failed
          msgv1  = CONV #( iv_file_path ).
    ENDIF.

  ENDMETHOD.

  METHOD is_windows.

    DATA lv_subrc TYPE sy-subrc.

    CALL FUNCTION 'HLP_OPERATING_SYSTEM_CHECK'
      IMPORTING
        returncode = lv_subrc.

    rv_result = COND #( WHEN lv_subrc EQ 0 THEN abap_true ELSE abap_false ).

  ENDMETHOD.

ENDCLASS.