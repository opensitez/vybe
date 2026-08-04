*> vybe-test: cobol/classes_inheritance_polymorphism/class_multiple_methods_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C4.
OBJECT.
METHOD-ID. A.
PROCEDURE DIVISION.
    DISPLAY "A".
END METHOD A.
METHOD-ID. B.
PROCEDURE DIVISION.
    DISPLAY "B".
END METHOD B.
END OBJECT.
END CLASS C4.

