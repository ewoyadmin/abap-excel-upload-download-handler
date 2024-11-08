"! Exception class for Excel handler
"! <p>
"! This class defines various exception scenarios that can occur during
"! Excel file operations, such as uploading, downloading, and data conversion.
"! </p>
CLASS zcx_excel_handler DEFINITION
  PUBLIC
  INHERITING FROM cx_dynamic_check
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    INTERFACES if_t100_message .
    INTERFACES if_t100_dyn_msg .

    CONSTANTS:
    "! <p class="shorttext synchronized">Column not found error</p>
      BEGIN OF column_not_found,
        msgid TYPE symsgid VALUE 'RSL_UI',
        msgno TYPE symsgno VALUE '283',
        attr1 TYPE scx_attrname VALUE 'MSGV1',
        attr2 TYPE scx_attrname VALUE '',
        attr3 TYPE scx_attrname VALUE '',
        attr4 TYPE scx_attrname VALUE '',
      END OF column_not_found .

    CONSTANTS:
    "! <p class="shorttext synchronized">Invalid value error</p>
      BEGIN OF invalid_value,
        msgid TYPE symsgid VALUE '/PRA/PN',
        msgno TYPE symsgno VALUE '008',
        attr1 TYPE scx_attrname VALUE 'MSGV1',
        attr2 TYPE scx_attrname VALUE 'MSGV2',
        attr3 TYPE scx_attrname VALUE '',
        attr4 TYPE scx_attrname VALUE '',
      END OF invalid_value .

    CONSTANTS:
    "! <p class="shorttext synchronized">File not found error</p>
      BEGIN OF file_not_found,
        msgid TYPE symsgid VALUE 'CH',
        msgno TYPE symsgno VALUE '132',
        attr1 TYPE scx_attrname VALUE 'MSGV1',
        attr2 TYPE scx_attrname VALUE 'MSGV2',
        attr3 TYPE scx_attrname VALUE '',
        attr4 TYPE scx_attrname VALUE '',
      END OF file_not_found .

    CONSTANTS:
    "! <p class="shorttext synchronized">Invalid date error</p>
      BEGIN OF invalid_date,
        msgid TYPE symsgid VALUE '00',
        msgno TYPE symsgno VALUE '302',
        attr1 TYPE scx_attrname VALUE 'MSGV1',
        attr2 TYPE scx_attrname VALUE '',
        attr3 TYPE scx_attrname VALUE '',
        attr4 TYPE scx_attrname VALUE '',
      END OF invalid_date .

    CONSTANTS:
    "! <p class="shorttext synchronized">Invalid delimiter error</p>
      BEGIN OF invalid_delimiter,
        msgid TYPE symsgid VALUE '/IBX/UI',
        msgno TYPE symsgno VALUE '016',
        attr1 TYPE scx_attrname VALUE '',
        attr2 TYPE scx_attrname VALUE '',
        attr3 TYPE scx_attrname VALUE '',
        attr4 TYPE scx_attrname VALUE '',
      END OF invalid_delimiter .

    CONSTANTS:
    "! <p class="shorttext synchronized">Other general error</p>
      BEGIN OF other_error,
        msgid TYPE symsgid VALUE '00',
        msgno TYPE symsgno VALUE '001',
        attr1 TYPE scx_attrname VALUE 'MSGV1',
        attr2 TYPE scx_attrname VALUE 'MSGV2',
        attr3 TYPE scx_attrname VALUE 'MSGV3',
        attr4 TYPE scx_attrname VALUE 'MSGV4',
      END OF other_error .

    CONSTANTS:
    "! <p class="shorttext synchronized">Conversion failed error</p>
      BEGIN OF conversion_failed,
        msgid TYPE symsgid VALUE 'OIUX1_TAX_MESSAGES',
        msgno TYPE symsgno VALUE '064',
        attr1 TYPE scx_attrname VALUE 'MSGV1',
        attr2 TYPE scx_attrname VALUE '',
        attr3 TYPE scx_attrname VALUE '',
        attr4 TYPE scx_attrname VALUE '',
      END OF conversion_failed .

    CONSTANTS:
    "! <p class="shorttext synchronized">File write error</p>
      BEGIN OF file_write_error,
        msgid TYPE symsgid VALUE '3O',
        msgno TYPE symsgno VALUE '254',
        attr1 TYPE scx_attrname VALUE 'MSGV1',
        attr2 TYPE scx_attrname VALUE '',
        attr3 TYPE scx_attrname VALUE '',
        attr4 TYPE scx_attrname VALUE '',
      END OF file_write_error .

    CONSTANTS:
    "! <p class="shorttext synchronized">Operation not supported error</p>
      BEGIN OF not_supported,
        msgid TYPE symsgid VALUE '/NFM/CA',
        msgno TYPE symsgno VALUE '023',
        attr1 TYPE scx_attrname VALUE '',
        attr2 TYPE scx_attrname VALUE '',
        attr3 TYPE scx_attrname VALUE '',
        attr4 TYPE scx_attrname VALUE '',
      END OF not_supported .

      CONSTANTS:
      "! <p class="shorttext synchronized">You selected data nodes with different structures</p>
        BEGIN OF different_structure,
          msgid TYPE symsgid VALUE 'QG_EVAL',
          msgno TYPE symsgno VALUE '033',
          attr1 TYPE scx_attrname VALUE '',
          attr2 TYPE scx_attrname VALUE '',
          attr3 TYPE scx_attrname VALUE '',
          attr4 TYPE scx_attrname VALUE '',
        END OF different_structure .
      
    "! <p class="shorttext synchronized">First message variable</p>
    DATA msgv1 TYPE msgv1 .
    "! <p class="shorttext synchronized">Second message variable</p>
    DATA msgv2 TYPE msgv2 .
    "! <p class="shorttext synchronized">Third message variable</p>
    DATA msgv3 TYPE msgv3 .
    "! <p class="shorttext synchronized">Fourth message variable</p>
    DATA msgv4 TYPE msgv4 .

    "! Constructor
    "! @parameter textid | Text ID
    "! @parameter previous | Previous exception
    "! @parameter msgv1 | Message variable 1
    "! @parameter msgv2 | Message variable 2
    "! @parameter msgv3 | Message variable 3
    "! @parameter msgv4 | Message variable 4
    METHODS constructor
      IMPORTING
        !textid   LIKE if_t100_message=>t100key OPTIONAL
        !previous LIKE previous OPTIONAL
        !msgv1    TYPE msgv1 OPTIONAL
        !msgv2    TYPE msgv2 OPTIONAL
        !msgv3    TYPE msgv3 OPTIONAL
        !msgv4    TYPE msgv4 OPTIONAL .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS zcx_excel_handler IMPLEMENTATION.
  METHOD constructor ##ADT_SUPPRESS_GENERATION.
    CALL METHOD super->constructor
      EXPORTING
        previous = previous.
    me->msgv1 = msgv1 .
    me->msgv2 = msgv2 .
    me->msgv3 = msgv3 .
    me->msgv4 = msgv4 .
    CLEAR me->textid.
    IF textid IS INITIAL.
      if_t100_message~t100key = if_t100_message=>default_textid.
    ELSE.
      if_t100_message~t100key = textid.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
