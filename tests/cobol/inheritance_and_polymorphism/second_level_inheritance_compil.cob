*> vybe-test: cobol/inheritance_and_polymorphism/second_level_inheritance_compiles
*> origin: languages/cobol/tests/cobol/test_inheritance_and_polymorphism.rs

IDENTIFICATION DIVISION.
CLASS-ID. SMART-CIRCLE INHERITS FROM CIRCLE.
OBJECT.
METHOD-ID. DESCRIBE.
PROCEDURE DIVISION.
    DISPLAY "SMART".
END METHOD DESCRIBE.
END OBJECT.
END CLASS SMART-CIRCLE.

