*> vybe-test: cobol/collating_sequence/collating_sequence_with_repository_compiles
*> origin: languages/cobol/tests/cobol/test_collating_sequence.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. COLL8.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    COLLATING SEQUENCE IS ALPHA8.
REPOSITORY.
    FUNCTION ALL INTRINSIC.
PROCEDURE DIVISION.
    STOP RUN.

