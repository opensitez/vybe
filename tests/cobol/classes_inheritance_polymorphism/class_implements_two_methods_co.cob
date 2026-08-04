*> vybe-test: cobol/classes_inheritance_polymorphism/class_implements_two_methods_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C6 IMPLEMENTS I2.
OBJECT.
METHOD-ID. M1.
PROCEDURE DIVISION.
    DISPLAY "1".
END METHOD M1.
METHOD-ID. M2.
PROCEDURE DIVISION.
    DISPLAY "2".
END METHOD M2.
END OBJECT.
END CLASS C6.

