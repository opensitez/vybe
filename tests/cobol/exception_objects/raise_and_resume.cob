*> vybe-test: cobol/exception_objects/raise_and_resume
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-x PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           ADD 1 TO ws-x
           RAISE EXCEPTION EC-SIZE-OVERFLOW
           ADD 1 TO ws-x
           DISPLAY ws-x
           STOP RUN.

