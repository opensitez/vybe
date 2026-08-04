*> vybe-test: cobol/class_clause/class_clause_if_negative_example_compiles
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS7.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS STAR-CLASS IS "*".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 CH PIC X VALUE "*".
PROCEDURE DIVISION.
    IF CH IS STAR-CLASS DISPLAY "Y" END-IF.
    STOP RUN.

