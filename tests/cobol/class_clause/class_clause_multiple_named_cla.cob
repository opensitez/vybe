*> vybe-test: cobol/class_clause/class_clause_multiple_named_classes_compiles
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS8.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS DIGITS IS "0" THRU "9".
    CLASS LETTERS IS "A" THRU "Z".
PROCEDURE DIVISION.
    STOP RUN.

