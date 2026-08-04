*> vybe-test: cobol/basic_classes_oop/class_method_with_parameters_compiles
*> origin: languages/cobol/tests/cobol/test_basic_classes_oop.rs

IDENTIFICATION DIVISION.
CLASS-ID. METER.
OBJECT.
METHOD-ID. SET-LIMIT.
PROCEDURE DIVISION USING WS-LIMIT.
    MOVE WS-LIMIT TO WS-LIMIT.
END METHOD SET-LIMIT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-LIMIT PIC 9(4).
END OBJECT.
END CLASS METER.

