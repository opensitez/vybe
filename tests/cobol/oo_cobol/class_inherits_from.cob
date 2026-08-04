*> vybe-test: cobol/oo_cobol/class_inherits_from
*> origin: languages/cobol/tests/cobol/test_oo_cobol.rs

IDENTIFICATION DIVISION.
CLASS-ID. DOG INHERITS FROM ANIMAL.
OBJECT.
METHOD-ID. SPEAK OVERRIDE.
PROCEDURE DIVISION.
    DISPLAY "Woof!".
END METHOD SPEAK.
END OBJECT.
END CLASS DOG.

