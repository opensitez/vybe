*> vybe-test: cobol/classes_inheritance_polymorphism/class_with_data_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. PERSON-C.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-N PIC X(20).
METHOD-ID. SET-N.
PROCEDURE DIVISION USING WS-IN.
    MOVE WS-IN TO WS-N.
END METHOD SET-N.
END OBJECT.
END CLASS PERSON-C.

