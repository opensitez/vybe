*> vybe-test: cobol/alphabet_clause/alphabet_clause_with_sort_collating_name_compiles
*> origin: languages/cobol/tests/cobol/test_alphabet_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA8.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET SORT-ALPHA IS STANDARD-1.
PROCEDURE DIVISION.
    STOP RUN.

