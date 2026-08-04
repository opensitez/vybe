*> vybe-test: cobol/special_names/decimal_point_comma_with_currency
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           DECIMAL-POINT IS COMMA
           CURRENCY SIGN IS "E".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-price PIC E9.999,99 VALUE 1234,56.
       PROCEDURE DIVISION.
           DISPLAY ws-price
           STOP RUN.

