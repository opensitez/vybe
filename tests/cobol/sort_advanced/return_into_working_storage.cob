*> vybe-test: cobol/sort_advanced/return_into_working_storage
*> origin: languages/cobol/tests/cobol/test_sort_advanced.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT sf ASSIGN TO "sf.tmp".
       DATA DIVISION.
       FILE SECTION.
       SD sf.
       01 srec.
           05 sk PIC X(5).
           05 sv PIC 99.
       WORKING-STORAGE SECTION.
       01 ws-buf  PIC X(7).
       01 ws-done PIC X VALUE "N".
       01 ws-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           SORT sf ON ASCENDING KEY sk
               USING "in.dat"
               OUTPUT PROCEDURE IS count-recs
           DISPLAY ws-count
           STOP RUN.
       count-recs SECTION.
           PERFORM UNTIL ws-done = "Y"
               RETURN sf INTO ws-buf
                   AT END MOVE "Y" TO ws-done
                   NOT AT END ADD 1 TO ws-count
               END-RETURN
           END-PERFORM.

