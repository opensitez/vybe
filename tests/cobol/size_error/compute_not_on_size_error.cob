*> vybe-test: cobol/size_error/compute_not_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 99  VALUE 0.
       01 ws-ok     PIC X   VALUE "N".
       PROCEDURE DIVISION.
           COMPUTE ws-result = 5 + 3
               NOT ON SIZE ERROR MOVE "Y" TO ws-ok
           END-COMPUTE
           DISPLAY ws-result
           DISPLAY ws-ok
           STOP RUN.

