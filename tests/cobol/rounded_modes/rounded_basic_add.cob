*> vybe-test: cobol/rounded_modes/rounded_basic_add
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9V9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED = 1.35
           DISPLAY ws-result
           STOP RUN.

