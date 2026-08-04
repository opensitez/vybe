*> vybe-test: cobol/declaratives/declaratives_use_for_debugging_specific
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-x PIC 99 VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       debug-x SECTION.
           USE FOR DEBUGGING ON ws-x.
           DISPLAY "ws-x changed to: " DEBUG-CONTENTS.
       END DECLARATIVES.
       main-para SECTION.
           MOVE 42 TO ws-x
           DISPLAY ws-x
           STOP RUN.

