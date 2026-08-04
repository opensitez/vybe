*> vybe-test: cobol/size_error/compute_overflow_takes_size_error_path_only
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 99 VALUE 0.
       01 ws-status PIC X VALUE SPACE.
       PROCEDURE DIVISION.
           COMPUTE ws-result = 999 * 999
               ON SIZE ERROR     MOVE "E" TO ws-status
               NOT ON SIZE ERROR MOVE "K" TO ws-status
           END-COMPUTE
           DISPLAY ws-status
           STOP RUN.

