*> vybe-test: cobol/classes_inheritance_polymorphism/class_property_get_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C1.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-V PIC 9(3).
METHOD-ID. GET-V PROPERTY GET.
PROCEDURE DIVISION RETURNING WS-R.
    MOVE WS-V TO WS-R.
END METHOD GET-V.
END OBJECT.
END CLASS C1.

