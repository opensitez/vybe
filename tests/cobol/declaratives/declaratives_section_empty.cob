*> vybe-test: cobol/declaratives/declaratives_section_empty
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-flag PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       error-section SECTION.
           USE AFTER STANDARD ERROR PROCEDURE.
           DISPLAY "error handler".
       END DECLARATIVES.
       main-logic SECTION.
           MOVE "Y" TO ws-flag
           DISPLAY ws-flag
           STOP RUN.

