*> vybe-test: cobol/collections_dictionaries/map_from_json_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 J PIC X(100).
PROCEDURE DIVISION.
    CALL "MAP-FROM-JSON" USING J.
    STOP RUN.

