*> vybe-test: cobol/special_names/decimal_point_comma_arithmetic
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           DECIMAL-POINT IS COMMA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a PIC 9V99 VALUE 1,50.
       01 ws-b PIC 9V99 VALUE 2,75.
       01 ws-c PIC 9V99 VALUE 0.
       PROCEDURE DIVISION.
           ADD ws-a ws-b GIVING ws-c
           DISPLAY ws-c
           STOP RUN.

