*> vybe-test: cobol/classes_inheritance_polymorphism/class_derived_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. DERIVED-C INHERITS FROM BASE-C.
OBJECT.
METHOD-ID. M1 OVERRIDE.
PROCEDURE DIVISION.
    DISPLAY "OV".
END METHOD M1.
END OBJECT.
END CLASS DERIVED-C.

