*> vybe-test: cobol/exception_objects/invoke_exception_method
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

       CLASS-ID. MyException INHERITS FROM EXCEPTION-OBJECT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-code PIC 9(4).
       METHOD-ID. GET-CODE.
       LINKAGE SECTION.
       01 lk-code PIC 9(4).
       PROCEDURE DIVISION RETURNING lk-code.
           MOVE ws-code TO lk-code
           GOBACK.
       END METHOD GET-CODE.
       END CLASS MyException.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS MyException AS "MyException".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-ex   OBJECT REFERENCE MyException.
       01 ws-code PIC 9(4).
       PROCEDURE DIVISION.
           INVOKE MyException NEW RETURNING ws-ex
           INVOKE ws-ex GET-CODE RETURNING ws-code
           DISPLAY ws-code
           STOP RUN.

