*> vybe-test: cobol/classes_inheritance_polymorphism/class_override_with_display_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C10 INHERITS FROM BASE-C.
OBJECT.
METHOD-ID. M1 OVERRIDE.
PROCEDURE DIVISION.
    DISPLAY "OV2".
END METHOD M1.
END OBJECT.
END CLASS C10.

