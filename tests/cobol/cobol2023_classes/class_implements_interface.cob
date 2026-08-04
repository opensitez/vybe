*> vybe-test: cobol/cobol2023_classes/class_implements_interface
*> origin: languages/cobol/tests/cobol/test_cobol2023_classes.rs

IDENTIFICATION DIVISION.
CLASS-ID. PRINTABLE-ITEM IMPLEMENTS PRINTABLE.
OBJECT.
METHOD-ID. PRINT-SELF.
PROCEDURE DIVISION.
    DISPLAY "Printing item".
END METHOD PRINT-SELF.
END OBJECT.
END CLASS PRINTABLE-ITEM.

