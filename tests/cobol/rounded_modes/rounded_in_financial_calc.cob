*> vybe-test: cobol/rounded_modes/rounded_in_financial_calc
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-principal  PIC 9(7)V99 VALUE 1000.00.
       01 ws-rate       PIC V9(4)   VALUE .0875.
       01 ws-interest   PIC 9(5)V99 VALUE 0.
       01 ws-total      PIC 9(7)V99 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-interest ROUNDED MODE NEAREST-EVEN
               = ws-principal * ws-rate
           COMPUTE ws-total = ws-principal + ws-interest
           DISPLAY ws-interest
           DISPLAY ws-total
           STOP RUN.

