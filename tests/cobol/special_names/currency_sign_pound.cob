*> vybe-test: cobol/special_names/currency_sign_pound
*> origin: languages/cobol/tests/cobol/test_special_names.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SPECIAL-NAMES.
           CURRENCY SIGN IS "L".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-amount PIC L9(5)V99 VALUE 500.00.
       PROCEDURE DIVISION.
           DISPLAY ws-amount
           STOP RUN.

