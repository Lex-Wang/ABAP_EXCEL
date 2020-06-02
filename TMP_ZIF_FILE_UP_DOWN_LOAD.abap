*"* components of interface ZIF_FILE_UP_DOWN_LOAD
interface ZIF_FILE_UP_DOWN_LOAD
  public .


  types:
    BEGIN OF ty_field_exc,
      field TYPE fieldname,
    END OF ty_field_exc .
  types:
    ty_t_field_exc TYPE STANDARD TABLE OF ty_field_exc .

  data _GV_FILENAME type STRING read-only .

  methods SET_TABLE_REF
    importing
      !PI_TABLE_REF type ref to DATA
      !PI_FILENAME type STRING .
  methods UPLOAD_FILE
    importing
      !PI_HD_INDX type I default 1
      !PI_HDTXT_INDX type I default 2
      !PI_CONTENT_INDEX type I default 3
      !PI_ERRMSG_FIELD type DD03D-FIELDNAME
    exceptions
      FAILED_TO_UPLOAD .
  methods DOWNLOAD_FILE
    importing
      value(PI_FIELD_EXC) type TY_T_FIELD_EXC optional
      !PI_CHARSET type CHAR20 default 'UTF-8'
    preferred parameter PI_FIELD_EXC
    exceptions
      CANNOT_DOWNLOAD .
  methods SET_FIELD_LABEL
    importing
      !PI_FIELD type DD03D-FIELDNAME
      !PI_FIELDTXT type DD03P-DDTEXT .
endinterface.
