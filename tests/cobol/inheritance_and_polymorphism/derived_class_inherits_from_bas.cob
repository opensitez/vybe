*> vybe-test: cobol/inheritance_and_polymorphism/derived_class_inherits_from_base_compiles
*> origin: languages/cobol/tests/cobol/test_inheritance_and_polymorphism.rs

IDENTIFICATION DIVISION.
CLASS-ID. CIRCLE INHERITS FROM SHAPE.
OBJECT.
METHOD-ID. AREA OVERRIDE.
PROCEDURE DIVISION RETURNING WS-RESULT.
    MOVE 314 TO WS-RESULT.
END METHOD AREA.
END OBJECT.
END CLASS CIRCLE.

