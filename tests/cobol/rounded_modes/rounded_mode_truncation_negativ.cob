*> vybe-test: cobol/rounded_modes/rounded_mode_truncation_negative
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC S9V9 VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE TRUNCATION = -2.79
           DISPLAY ws-result
           STOP RUN.

