*> vybe-test: cobol/exception_objects/resume_after_exception
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-step  PIC 9 VALUE 0.
       01 ws-total PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           MOVE 1 TO ws-step
           ADD ws-step TO ws-total
           RAISE EXCEPTION EC-PROGRAM-RECURSIVE-CALL
           MOVE 2 TO ws-step
           ADD ws-step TO ws-total
           DISPLAY ws-total
           STOP RUN.

