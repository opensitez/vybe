*> vybe-test: cobol/size_error/compute_both_branches_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-target PIC 99 VALUE 0.
       01 ws-status PIC X  VALUE SPACE.
       PROCEDURE DIVISION.
           COMPUTE ws-target = 200 + 300
               ON SIZE ERROR     MOVE "E" TO ws-status
               NOT ON SIZE ERROR MOVE "K" TO ws-status
           END-COMPUTE
           DISPLAY ws-status
           STOP RUN.

