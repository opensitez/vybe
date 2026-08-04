*> vybe-test: cobol/classes_inheritance_polymorphism/class_factory_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. UTIL.
FACTORY.
METHOD-ID. BUILD.
PROCEDURE DIVISION RETURNING WS-OBJ.
    DISPLAY "B".
END METHOD BUILD.
END FACTORY.
END CLASS UTIL.

