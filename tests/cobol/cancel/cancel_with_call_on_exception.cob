*> vybe-test: cobol/cancel/cancel_with_call_on_exception
*> origin: languages/cobol/tests/cobol/test_cancel.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-loaded PIC X VALUE "N".
       PROCEDURE DIVISION.
           CALL "plugin-module"
               ON EXCEPTION
                   DISPLAY "load failed"
                   GO TO end-prog
               NOT ON EXCEPTION
                   MOVE "Y" TO ws-loaded
           END-CALL
           IF ws-loaded = "Y"
               CANCEL "plugin-module"
           END-IF
       end-prog.
           STOP RUN.

