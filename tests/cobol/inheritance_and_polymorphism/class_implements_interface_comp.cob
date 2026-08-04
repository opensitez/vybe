*> vybe-test: cobol/inheritance_and_polymorphism/class_implements_interface_compiles
*> origin: languages/cobol/tests/cobol/test_inheritance_and_polymorphism.rs

IDENTIFICATION DIVISION.
CLASS-ID. REPORT IMPLEMENTS PRINTABLE.
OBJECT.
METHOD-ID. PRINT-SELF.
PROCEDURE DIVISION.
    DISPLAY "REPORT".
END METHOD PRINT-SELF.
END OBJECT.
END CLASS REPORT.

