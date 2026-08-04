*> vybe-test: cobol/size_error/subtract_not_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 99 VALUE 20.
       01 ws-ok  PIC X VALUE "N".
       PROCEDURE DIVISION.
           SUBTRACT 5 FROM ws-val
               NOT ON SIZE ERROR MOVE "Y" TO ws-ok
           END-SUBTRACT
           DISPLAY ws-ok
           STOP RUN.

