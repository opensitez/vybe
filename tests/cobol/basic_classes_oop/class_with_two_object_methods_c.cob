*> vybe-test: cobol/basic_classes_oop/class_with_two_object_methods_compiles
*> origin: languages/cobol/tests/cobol/test_basic_classes_oop.rs

IDENTIFICATION DIVISION.
CLASS-ID. COUNTER.
OBJECT.
METHOD-ID. INC.
PROCEDURE DIVISION.
    DISPLAY "INC".
END METHOD INC.
METHOD-ID. GET.
PROCEDURE DIVISION RETURNING WS-V.
    MOVE 1 TO WS-V.
END METHOD GET.
END OBJECT.
END CLASS COUNTER.

