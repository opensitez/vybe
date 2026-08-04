*> vybe-test: cobol/classes_inheritance_polymorphism/class_invoke_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_classes_inheritance_polymorphism.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. P.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 O USAGE POINTER.
01 R PIC 9(3).
PROCEDURE DIVISION.
    INVOKE O CODE RETURNING R.
    STOP RUN.

