*> vybe-test: cobol/level_66_78/renames_single_field
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 payment-rec.
           05 pay-amount    PIC 9(8)V99.
           05 pay-currency  PIC XXX.
       66 pay-curr-alias RENAMES pay-currency.
       PROCEDURE DIVISION.
           MOVE "USD" TO pay-currency
           DISPLAY pay-curr-alias
           STOP RUN.

