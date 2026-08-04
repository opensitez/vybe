*> vybe-test: cobol/exception_objects/exception_hierarchy_three_levels
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

       CLASS-ID. AppBaseEx INHERITS FROM EXCEPTION-OBJECT.
       END CLASS AppBaseEx.

       CLASS-ID. IOException INHERITS FROM AppBaseEx.
       END CLASS IOException.

       CLASS-ID. FileNotFoundException INHERITS FROM IOException.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-filename PIC X(200).
       METHOD-ID. SET-FILENAME.
       LINKAGE SECTION.
       01 lk-fn PIC X(200).
       PROCEDURE DIVISION USING lk-fn.
           MOVE lk-fn TO ws-filename
           GOBACK.
       END METHOD SET-FILENAME.
       END CLASS FileNotFoundException.
    STOP RUN.

