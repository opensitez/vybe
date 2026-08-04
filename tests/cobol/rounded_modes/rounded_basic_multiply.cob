*> vybe-test: cobol/rounded_modes/rounded_basic_multiply
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           MULTIPLY 3 BY 3.7 GIVING ws-result ROUNDED
           DISPLAY ws-result
           STOP RUN.

