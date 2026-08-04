*> vybe-test: cobol/collections_dictionaries/keyed_exists_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC X(5) VALUE "A".
01 E PIC 9 VALUE 0.
PROCEDURE DIVISION.
    CALL "MAP-EXISTS" USING K E.
    STOP RUN.

