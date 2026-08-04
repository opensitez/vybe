*> vybe-test: cobol/declaratives/declaratives_with_working_storage_access
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-error-code  PIC 99  VALUE 0.
       01 ws-error-msg   PIC X(40) VALUE SPACES.
       01 ws-error-count PIC 999  VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       error-handler SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           ADD 1 TO ws-error-count
           MOVE ws-error-count TO ws-error-code
           STRING "Error #" DELIMITED SIZE
                  ws-error-code DELIMITED SIZE
                  INTO ws-error-msg.
       END DECLARATIVES.
       main-section SECTION.
           DISPLAY ws-error-count
           STOP RUN.

