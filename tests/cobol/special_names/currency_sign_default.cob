*> vybe-test: cobol/special_names/currency_sign_default
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CURRENCY SIGN IS "$".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-amount PIC $9,999.99 VALUE 1234.56.
       PROCEDURE DIVISION.
           DISPLAY ws-amount
           STOP RUN.

