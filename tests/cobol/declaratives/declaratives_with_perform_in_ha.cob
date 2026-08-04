*> vybe-test: cobol/declaratives/declaratives_with_perform_in_handler
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-err-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       err-handler SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           PERFORM log-error.
       END DECLARATIVES.
       main-section SECTION.
           DISPLAY ws-err-count
           STOP RUN.
       log-error.
           ADD 1 TO ws-err-count
           DISPLAY "Error logged: " ws-err-count.

