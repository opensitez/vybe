*> vybe-test: cobol/exception_objects/exception_class_with_constructor_compiles
*> origin: languages/cobol/tests/cobol/test_exception_objects.rs

IDENTIFICATION DIVISION.
CLASS-ID. TimeoutException INHERITS FROM EXCEPTION-OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TIMEOUT PIC 9(9) COMP.
METHOD-ID. NEW.
LINKAGE SECTION.
01 LK-TIMEOUT PIC 9(9) COMP.
PROCEDURE DIVISION USING LK-TIMEOUT.
    MOVE LK-TIMEOUT TO WS-TIMEOUT
    GOBACK.
END METHOD NEW.
END CLASS TimeoutException.

