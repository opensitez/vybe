*> vybe-test: cobol/inheritance_and_polymorphism/deep_inheritance_override_chain_compiles
*> origin: languages/cobol/tests/cobol/test_inheritance_and_polymorphism.rs

IDENTIFICATION DIVISION.
CLASS-ID. BASE-T.
OBJECT.
METHOD-ID. NAME.
PROCEDURE DIVISION.
    DISPLAY "BASE".
END METHOD NAME.
END OBJECT.
END CLASS BASE-T.

IDENTIFICATION DIVISION.
CLASS-ID. MID-T INHERITS FROM BASE-T.
OBJECT.
METHOD-ID. NAME OVERRIDE.
PROCEDURE DIVISION.
    DISPLAY "MID".
END METHOD NAME.
END OBJECT.
END CLASS MID-T.

