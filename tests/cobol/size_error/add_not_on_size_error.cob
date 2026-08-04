*> vybe-test: cobol/size_error/add_not_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val     PIC 99 VALUE 5.
       01 ws-ok-flag PIC X VALUE "N".
       PROCEDURE DIVISION.
           ADD 3 TO ws-val
               NOT ON SIZE ERROR MOVE "Y" TO ws-ok-flag
           END-ADD
           DISPLAY ws-ok-flag
           STOP RUN.

