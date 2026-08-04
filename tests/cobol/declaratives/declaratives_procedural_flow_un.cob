*> vybe-test: cobol/declaratives/declaratives_procedural_flow_unaffected
*> origin: languages/cobol/tests/cobol/test_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9(5) VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       err-sec SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           DISPLAY "error".
       END DECLARATIVES.
       main-section SECTION.
           PERFORM VARYING ws-result FROM 1 BY 1
               UNTIL ws-result > 5
               CONTINUE
           END-PERFORM
           DISPLAY ws-result
           STOP RUN.

