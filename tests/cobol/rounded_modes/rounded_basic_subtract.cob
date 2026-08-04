*> vybe-test: cobol/rounded_modes/rounded_basic_subtract
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a      PIC 9V99 VALUE 5.67.
       01 ws-b      PIC 9V99 VALUE 2.34.
       01 ws-result PIC 9V9  VALUE 0.
       PROCEDURE DIVISION.
           SUBTRACT ws-b FROM ws-a GIVING ws-result ROUNDED
           DISPLAY ws-result
           STOP RUN.

