*> vybe-test: cobol/figurative_constants/all_in_compare
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(10) VALUE ALL "*".
       PROCEDURE DIVISION.
           IF ws-field = ALL "*"
               DISPLAY "all stars"
           ELSE
               DISPLAY "not all stars"
           END-IF
           STOP RUN.

