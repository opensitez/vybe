*> vybe-test: cobol/class_clause/class_clause_lowercase_range_compiles
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS4.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS LOWER-CLASS IS "a" THRU "z".
PROCEDURE DIVISION.
    STOP RUN.

