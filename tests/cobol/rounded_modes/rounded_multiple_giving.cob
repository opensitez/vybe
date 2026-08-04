*> vybe-test: cobol/rounded_modes/rounded_multiple_giving
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a   PIC 9V9 VALUE 3.7.
       01 ws-b   PIC 9V9 VALUE 3.7.
       01 ws-c   PIC 9V9 VALUE 3.7.
       PROCEDURE DIVISION.
           ADD 1.25 TO ws-a ROUNDED
           ADD 1.25 TO ws-b ROUNDED MODE TRUNCATION
           ADD 1.25 TO ws-c ROUNDED MODE NEAREST-EVEN
           DISPLAY ws-a
           DISPLAY ws-b
           DISPLAY ws-c
           STOP RUN.

