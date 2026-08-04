*> vybe-test: cobol/size_error/divide_giving_remainder_on_size_error
*> origin: languages/cobol/tests/cobol/test_size_error.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-a      PIC 9   VALUE 7.
       01 ws-b      PIC 9   VALUE 2.
       01 ws-q      PIC 9   VALUE 0.
       01 ws-r      PIC 9   VALUE 0.
       01 ws-err    PIC X   VALUE "N".
       PROCEDURE DIVISION.
           DIVIDE ws-b INTO ws-a
               GIVING ws-q REMAINDER ws-r
               NOT ON SIZE ERROR MOVE "N" TO ws-err
           END-DIVIDE
           DISPLAY ws-q
           DISPLAY ws-r
           STOP RUN.

