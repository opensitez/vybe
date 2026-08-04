*> vybe-test: cobol/figurative_constants/low_value_singular
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-nul PIC X.
       PROCEDURE DIVISION.
           MOVE LOW-VALUE TO ws-nul
           IF ws-nul = LOW-VALUE
               DISPLAY "is low value"
           END-IF
           STOP RUN.

