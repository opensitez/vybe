*> vybe-test: cobol/exception_objects/raise_exception_class
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

       CLASS-ID. AppException INHERITS FROM EXCEPTION-OBJECT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-message PIC X(100).
       METHOD-ID. GET-MESSAGE.
       LINKAGE SECTION.
       01 lk-msg PIC X(100).
       PROCEDURE DIVISION RETURNING lk-msg.
           MOVE ws-message TO lk-msg
           GOBACK.
       END METHOD GET-MESSAGE.
       METHOD-ID. SET-MESSAGE.
       LINKAGE SECTION.
       01 lk-msg PIC X(100).
       PROCEDURE DIVISION USING lk-msg.
           MOVE lk-msg TO ws-message
           GOBACK.
       END METHOD SET-MESSAGE.
       END CLASS AppException.
    STOP RUN.

