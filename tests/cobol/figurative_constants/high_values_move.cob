*> vybe-test: cobol/figurative_constants/high_values_move
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-key PIC X(10).
       PROCEDURE DIVISION.
           MOVE HIGH-VALUES TO ws-key
           DISPLAY "high values set"
           STOP RUN.

