*> vybe-test: cobol/classes_inheritance_polymorphism/class_method_returning_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C3.
OBJECT.
METHOD-ID. CODE.
PROCEDURE DIVISION RETURNING WS-R.
    MOVE 7 TO WS-R.
END METHOD CODE.
END OBJECT.
END CLASS C3.

