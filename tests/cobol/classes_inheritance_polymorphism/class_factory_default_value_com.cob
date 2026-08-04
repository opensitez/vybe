*> vybe-test: cobol/classes_inheritance_polymorphism/class_factory_default_value_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C11.
FACTORY.
METHOD-ID. NEW.
PROCEDURE DIVISION RETURNING WS-VAL.
    MOVE "OKAY" TO WS-VAL.
END METHOD NEW.
END FACTORY.
END CLASS C11.

