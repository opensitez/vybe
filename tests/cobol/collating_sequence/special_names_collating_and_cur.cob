*> vybe-test: cobol/collating_sequence/special_names_collating_and_currency_compiles
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL7.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    COLLATING SEQUENCE IS ALPHA7.
    CURRENCY SIGN IS "$".
PROCEDURE DIVISION.
    STOP RUN.

