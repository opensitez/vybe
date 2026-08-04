*> vybe-test: cobol/exception_objects/exception_hierarchy_two_levels
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

       CLASS-ID. BaseException INHERITS FROM EXCEPTION-OBJECT.
       END CLASS BaseException.

       CLASS-ID. DatabaseException INHERITS FROM BaseException.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-sql-code PIC S9(9) COMP.
       METHOD-ID. SET-SQL-CODE.
       LINKAGE SECTION.
       01 lk-code PIC S9(9) COMP.
       PROCEDURE DIVISION USING lk-code.
           MOVE lk-code TO ws-sql-code
           GOBACK.
       END METHOD SET-SQL-CODE.
       END CLASS DatabaseException.
    STOP RUN.

