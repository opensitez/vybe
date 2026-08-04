*> vybe-test: cobol/classes_inheritance_polymorphism/class_second_level_inherit_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C5 INHERITS FROM DERIVED-C.
OBJECT.
METHOD-ID. X.
PROCEDURE DIVISION.
    DISPLAY "X".
END METHOD X.
END OBJECT.
END CLASS C5.

