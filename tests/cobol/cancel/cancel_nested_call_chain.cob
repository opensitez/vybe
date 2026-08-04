*> vybe-test: cobol/cancel/cancel_nested_call_chain
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           CALL "level-1"
           CANCEL "level-1"
           CALL "level-2"
           CANCEL "level-2"
           CALL "level-3"
           CANCEL "level-3"
           STOP RUN.

