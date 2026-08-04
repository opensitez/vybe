*> vybe-test: cobol/inheritance_and_polymorphism/interface_extends_other_interface_compiles
*> origin: languages/cobol/tests/cobol/test_inheritance_and_polymorphism.rs

IDENTIFICATION DIVISION.
INTERFACE-ID. I-PRINT.
METHOD-ID. PRINT.
PROCEDURE DIVISION.
END METHOD PRINT.
END INTERFACE I-PRINT.

INTERFACE-ID. I-ADV-PRINT INHERITS FROM I-PRINT.
METHOD-ID. FORMAT.
PROCEDURE DIVISION.
END METHOD FORMAT.
END INTERFACE I-ADV-PRINT.

