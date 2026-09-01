*> vybe-test: cobol/collections_dictionaries/keyed_iterate_pattern_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 1.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5
        CALL "MAP-NEXT" USING I
    END-PERFORM.
    STOP RUN.

