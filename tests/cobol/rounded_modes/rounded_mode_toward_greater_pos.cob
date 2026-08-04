*> vybe-test: cobol/rounded_modes/rounded_mode_toward_greater_positive
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE TOWARD-GREATER = 2.1
           DISPLAY ws-result
           STOP RUN.

