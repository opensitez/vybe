*> vybe-test: cobol/classes_inheritance_polymorphism/class_static_like_factory_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C7.
FACTORY.
METHOD-ID. NEW-OBJ.
PROCEDURE DIVISION RETURNING WS-O.
    DISPLAY "N".
END METHOD NEW-OBJ.
END FACTORY.
END CLASS C7.

