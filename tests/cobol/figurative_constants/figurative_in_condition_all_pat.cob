*> vybe-test: cobol/figurative_constants/figurative_in_condition_all_paths
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field PIC X(5).
       PROCEDURE DIVISION.
           MOVE "HELLO" TO ws-field
           EVALUATE TRUE
               WHEN ws-field = SPACES     DISPLAY "spaces"
               WHEN ws-field = HIGH-VALUES DISPLAY "high"
               WHEN ws-field = LOW-VALUES  DISPLAY "low"
               WHEN ws-field = ZEROS       DISPLAY "zeros"
               WHEN ws-field = ALL "*"     DISPLAY "stars"
               WHEN OTHER                  DISPLAY ws-field
           END-EVALUATE
           STOP RUN.

