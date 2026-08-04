*> vybe-test: cobol/figurative_constants/high_value_singular
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-sentinel PIC X.
       PROCEDURE DIVISION.
           MOVE HIGH-VALUE TO ws-sentinel
           IF ws-sentinel = HIGH-VALUE
               DISPLAY "is high value"
           END-IF
           STOP RUN.

