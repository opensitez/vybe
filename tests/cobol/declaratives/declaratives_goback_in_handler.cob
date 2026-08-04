*> vybe-test: cobol/declaratives/declaratives_goback_in_handler
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-handled PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       fatal-error SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           MOVE "Y" TO ws-handled
           DISPLAY "Fatal error - handled".
       END DECLARATIVES.
       main-section SECTION.
           DISPLAY ws-handled
           STOP RUN.

