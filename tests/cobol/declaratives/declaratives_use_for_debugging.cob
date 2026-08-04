*> vybe-test: cobol/declaratives/declaratives_use_for_debugging
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-counter PIC 999 VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       debug-section SECTION.
           USE FOR DEBUGGING ON ALL PROCEDURES.
           DISPLAY "Debugging: " DEBUG-NAME.
       END DECLARATIVES.
       main-section SECTION.
           ADD 1 TO ws-counter
           DISPLAY ws-counter
           STOP RUN.

