*> vybe-test: cobol/class_clause/special_names_class_clause_compiles
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS1.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS DIGIT-CLASS IS "0" THRU "9".
PROCEDURE DIVISION.
    STOP RUN.

