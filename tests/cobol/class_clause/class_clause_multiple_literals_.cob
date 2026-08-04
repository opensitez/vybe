*> vybe-test: cobol/class_clause/class_clause_multiple_literals_compiles
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS3.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS HEX-CLASS IS "A" THRU "F" "0" THRU "9".
PROCEDURE DIVISION.
    STOP RUN.

