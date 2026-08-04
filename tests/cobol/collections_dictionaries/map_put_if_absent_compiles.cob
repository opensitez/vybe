*> vybe-test: cobol/collections_dictionaries/map_put_if_absent_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC X(5) VALUE "B".
01 V PIC X(10) VALUE "X".
PROCEDURE DIVISION.
    CALL "MAP-PUT-IF-ABSENT" USING K V.
    STOP RUN.

