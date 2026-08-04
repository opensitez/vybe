*> vybe-test: cobol/basic_classes_oop/class_with_object_section_compiles
*> origin: languages/cobol/tests/cobol/test_basic_classes_oop.rs

IDENTIFICATION DIVISION.
CLASS-ID. LOGGER.
OBJECT.
METHOD-ID. LOG-MSG.
PROCEDURE DIVISION.
    DISPLAY "LOG".
END METHOD LOG-MSG.
END OBJECT.
END CLASS LOGGER.

