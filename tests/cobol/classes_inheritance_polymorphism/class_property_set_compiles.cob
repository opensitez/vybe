*> vybe-test: cobol/classes_inheritance_polymorphism/class_property_set_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C2.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-V PIC 9(3).
METHOD-ID. SET-V PROPERTY SET.
PROCEDURE DIVISION USING WS-I.
    MOVE WS-I TO WS-V.
END METHOD SET-V.
END OBJECT.
END CLASS C2.

