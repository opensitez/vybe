*> vybe-test: cobol/exception_objects/exception_in_subprogram
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. main-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC X VALUE "N".
       PROCEDURE DIVISION.
           CALL "sub-prog" USING ws-result
               ON EXCEPTION MOVE "E" TO ws-result
           END-CALL
           DISPLAY ws-result
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. sub-prog.
       DATA DIVISION.
       LINKAGE SECTION.
       01 lk-result PIC X.
       PROCEDURE DIVISION USING lk-result.
           MOVE "Y" TO lk-result
           GOBACK.

