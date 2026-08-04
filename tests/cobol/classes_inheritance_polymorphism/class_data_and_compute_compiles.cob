*> vybe-test: cobol/classes_inheritance_polymorphism/class_data_and_compute_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
CLASS-ID. C9.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 2.
01 B PIC 9 VALUE 3.
METHOD-ID. SUM.
PROCEDURE DIVISION RETURNING R.
    COMPUTE R = A + B.
END METHOD SUM.
END OBJECT.
END CLASS C9.

