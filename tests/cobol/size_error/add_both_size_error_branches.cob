*> vybe-test: cobol/size_error/add_both_size_error_branches
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-num    PIC 9 VALUE 8.
       01 ws-status PIC X VALUE SPACE.
       PROCEDURE DIVISION.
           ADD 5 TO ws-num
               ON SIZE ERROR     MOVE "O" TO ws-status
               NOT ON SIZE ERROR MOVE "K" TO ws-status
           END-ADD
           DISPLAY ws-status
           STOP RUN.

