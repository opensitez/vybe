*> vybe-test: cobol/category_file_io_delete_invalid_key/test_delete_not_invalid_key
*> origin: languages/cobol/tests/cobol/test_category_file_io_delete_invalid_key.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL. SELECT F ASSIGN TO 'a' ORGANIZATION IS INDEXED ACCESS IS RANDOM RECORD KEY IS K. DATA DIVISION. FILE SECTION. FD F. 01 R. 05 K PIC X. PROCEDURE DIVISION. DELETE F RECORD NOT INVALID KEY DISPLAY 'OK'. STOP RUN.

