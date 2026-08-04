*> vybe-test: cobol/exception_objects/custom_exception_class
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

       CLASS-ID. ValidationException INHERITS FROM EXCEPTION-OBJECT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field-name PIC X(30).
       01 ws-error-msg  PIC X(100).
       METHOD-ID. INITIALIZE-EX.
       LINKAGE SECTION.
       01 lk-field PIC X(30).
       01 lk-msg   PIC X(100).
       PROCEDURE DIVISION USING lk-field lk-msg.
           MOVE lk-field TO ws-field-name
           MOVE lk-msg   TO ws-error-msg
           GOBACK.
       END METHOD INITIALIZE-EX.
       END CLASS ValidationException.
    STOP RUN.

