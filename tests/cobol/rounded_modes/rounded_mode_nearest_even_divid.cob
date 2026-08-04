*> vybe-test: cobol/rounded_modes/rounded_mode_nearest_even_divide
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9V99 VALUE 0.
       PROCEDURE DIVISION.
           DIVIDE 3 INTO 1 GIVING ws-result ROUNDED MODE NEAREST-EVEN
           DISPLAY ws-result
           STOP RUN.

