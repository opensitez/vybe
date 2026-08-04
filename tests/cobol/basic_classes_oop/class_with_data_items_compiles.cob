*> vybe-test: cobol/basic_classes_oop/class_with_data_items_compiles
*> origin: languages/cobol/tests/cobol/test_basic_classes_oop.rs

IDENTIFICATION DIVISION.
CLASS-ID. PERSON.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(20).
01 WS-AGE PIC 9(3).
METHOD-ID. SET-NAME.
PROCEDURE DIVISION USING WS-IN.
    MOVE WS-IN TO WS-NAME.
END METHOD SET-NAME.
END OBJECT.
END CLASS PERSON.

