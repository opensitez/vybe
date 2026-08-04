*> vybe-test: cobol/level_66_78/level_78_used_in_compute
*> origin: languages/cobol/tests/cobol/test_level_66_78.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       78 TAX-RATE       VALUE 0.085.
       78 DISCOUNT-RATE  VALUE 0.10.
       01 ws-price       PIC 9(5)V99 VALUE 100.00.
       01 ws-tax         PIC 9(5)V99.
       01 ws-discount    PIC 9(5)V99.
       PROCEDURE DIVISION.
           COMPUTE ws-tax     = ws-price * TAX-RATE
           COMPUTE ws-discount = ws-price * DISCOUNT-RATE
           DISPLAY ws-tax
           DISPLAY ws-discount
           STOP RUN.

