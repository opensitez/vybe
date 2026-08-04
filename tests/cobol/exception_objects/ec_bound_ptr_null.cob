*> vybe-test: cobol/exception_objects/ec_bound_ptr_null
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-ptr USAGE POINTER VALUE NULL.
       PROCEDURE DIVISION.
           IF ws-ptr = NULL
               RAISE EXCEPTION EC-BOUND-PTR-NULL
           END-IF
           DISPLAY "checked"
           STOP RUN.

