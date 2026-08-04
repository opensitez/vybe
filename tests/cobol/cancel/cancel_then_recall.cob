*> vybe-test: cobol/cancel/cancel_then_recall
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9(5) VALUE 0.
       PROCEDURE DIVISION.
           CALL "counter-sub" USING ws-result
           DISPLAY ws-result
           CANCEL "counter-sub"
           MOVE 0 TO ws-result
           CALL "counter-sub" USING ws-result
           DISPLAY ws-result
           STOP RUN.

