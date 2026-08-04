*> vybe-test: cobol/alphabet_clause/alphabet_clause_multiple_ranges_compiles
*> origin: languages/cobol/tests/cobol/test_alphabet_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ALPHA7.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    ALPHABET HEX-SET IS "0" THRU "9" "A" THRU "F".
PROCEDURE DIVISION.
    STOP RUN.

