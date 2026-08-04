*> vybe-test: cobol/class_clause/class_clause_zero_and_space_mix_compiles
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS16.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS ZERO-CLASS IS ZERO.
    CLASS SPACE-CLASS IS SPACE.
PROCEDURE DIVISION.
    STOP RUN.

