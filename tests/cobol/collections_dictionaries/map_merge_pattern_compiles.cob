*> vybe-test: cobol/collections_dictionaries/map_merge_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    CALL "MAP-MERGE".
    STOP RUN.

