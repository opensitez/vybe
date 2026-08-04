*> vybe-test: cobol/basic_classes_oop/class_method_returning_compiles
*> origin: languages/cobol/tests/cobol/test_basic_classes_oop.rs

IDENTIFICATION DIVISION.
CLASS-ID. METRICS.
OBJECT.
METHOD-ID. GET-CODE.
PROCEDURE DIVISION RETURNING WS-RESULT.
    MOVE 7 TO WS-RESULT.
END METHOD GET-CODE.
END OBJECT.
END CLASS METRICS.

