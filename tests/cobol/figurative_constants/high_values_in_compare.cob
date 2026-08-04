*> vybe-test: cobol/figurative_constants/high_values_in_compare
*> origin: languages/cobol/tests/cobol/test_figurative_constants.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-key   PIC X(5) VALUE "ZZZZZ".
       01 ws-limit PIC X(5).
       PROCEDURE DIVISION.
           MOVE HIGH-VALUES TO ws-limit
           IF ws-key < ws-limit
               DISPLAY "key is below max"
           END-IF
           STOP RUN.

