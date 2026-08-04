*> vybe-test: cobol/class_clause/class_clause_alphanumeric_set_compiles
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS6.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS ID-CLASS IS "A" THRU "Z" "0" THRU "9" "-".
PROCEDURE DIVISION.
    STOP RUN.

