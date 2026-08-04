*> vybe-test: cobol/rounded_modes/rounded_mode_truncation_divide
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-q PIC 9V99 VALUE 0.
       PROCEDURE DIVISION.
           DIVIDE 6 INTO 10 GIVING ws-q ROUNDED MODE TRUNCATION
           DISPLAY ws-q
           STOP RUN.

