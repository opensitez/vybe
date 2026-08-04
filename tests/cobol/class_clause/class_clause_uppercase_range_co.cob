*> vybe-test: cobol/class_clause/class_clause_uppercase_range_compiles
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS5.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS UPPER-CLASS IS "A" THRU "Z".
PROCEDURE DIVISION.
    STOP RUN.

