*> vybe-test: cobol/exception_objects/ec_program_arg_omitted
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-err PIC X VALUE "N".
       PROCEDURE DIVISION.
           IF ws-err = "N"
               RAISE EXCEPTION EC-PROGRAM-ARG-OMITTED
           END-IF
           STOP RUN.

