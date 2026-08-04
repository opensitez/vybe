*> vybe-test: cobol/special_names/decimal_point_is_comma
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           DECIMAL-POINT IS COMMA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 9.999 VALUE 3,14159.
       PROCEDURE DIVISION.
           DISPLAY ws-val
           STOP RUN.

