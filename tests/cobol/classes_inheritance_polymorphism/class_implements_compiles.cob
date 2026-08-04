*> vybe-test: cobol/classes_inheritance_polymorphism/class_implements_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. DOC IMPLEMENTS IPRINT.
OBJECT.
METHOD-ID. PRINT-SELF.
PROCEDURE DIVISION.
    DISPLAY "DOC".
END METHOD PRINT-SELF.
END OBJECT.
END CLASS DOC.

