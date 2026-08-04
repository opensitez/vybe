*> vybe-test: cobol/collating_sequence/special_names_collating_sequence_compiles
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL1.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    COLLATING SEQUENCE IS ALPHA1.
PROCEDURE DIVISION.
    STOP RUN.

