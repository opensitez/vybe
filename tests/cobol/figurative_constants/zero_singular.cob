*> vybe-test: cobol/figurative_constants/zero_singular
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-counter PIC 99 VALUE 5.
       PROCEDURE DIVISION.
           IF ws-counter = ZERO
               DISPLAY "zero"
           ELSE
               DISPLAY "non-zero"
           END-IF
           STOP RUN.

