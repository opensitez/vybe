*> vybe-test: cobol/rounded_modes/rounded_mode_prohibited_exact
*> origin: languages/cobol/tests/cobol/test_rounded_modes.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9V9 VALUE 0.
       01 ws-err    PIC X   VALUE "N".
       PROCEDURE DIVISION.
           COMPUTE ws-result ROUNDED MODE PROHIBITED = 2.5
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-COMPUTE
           DISPLAY ws-err
           STOP RUN.

