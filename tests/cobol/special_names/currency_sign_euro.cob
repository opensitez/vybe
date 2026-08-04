*> vybe-test: cobol/special_names/currency_sign_euro
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CURRENCY SIGN IS "E".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-price PIC E9(6)V99 VALUE 1000.00.
       PROCEDURE DIVISION.
           DISPLAY ws-price
           STOP RUN.

