*> vybe-test: cobol/size_error/add_giving_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a   PIC 999 VALUE 999.
       01 ws-b   PIC 999 VALUE 1.
       01 ws-res PIC 999 VALUE 0.
       01 ws-err PIC X   VALUE "N".
       PROCEDURE DIVISION.
           ADD ws-a ws-b GIVING ws-res
               ON SIZE ERROR MOVE "Y" TO ws-err
           END-ADD
           DISPLAY ws-err
           STOP RUN.

