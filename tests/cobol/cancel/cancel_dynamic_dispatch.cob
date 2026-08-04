*> vybe-test: cobol/cancel/cancel_dynamic_dispatch
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-handler PIC X(30).
       01 ws-mode    PIC X(10) VALUE "fast".
       PROCEDURE DIVISION.
           IF ws-mode = "fast"
               MOVE "fast-handler" TO ws-handler
           ELSE
               MOVE "slow-handler" TO ws-handler
           END-IF
           CALL ws-handler
           CANCEL ws-handler
           DISPLAY "dispatched"
           STOP RUN.

