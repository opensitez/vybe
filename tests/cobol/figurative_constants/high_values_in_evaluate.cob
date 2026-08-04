*> vybe-test: cobol/figurative_constants/high_values_in_evaluate
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-status PIC X(5).
       PROCEDURE DIVISION.
           MOVE HIGH-VALUES TO ws-status
           EVALUATE ws-status
               WHEN HIGH-VALUES DISPLAY "end of data"
               WHEN LOW-VALUES  DISPLAY "start of data"
               WHEN OTHER       DISPLAY "data: " ws-status
           END-EVALUATE
           STOP RUN.

