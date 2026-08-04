*> vybe-test: cobol/collections_dictionaries/keyed_size_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(5).
PROCEDURE DIVISION.
    CALL "MAP-SIZE" USING N.
    STOP RUN.

