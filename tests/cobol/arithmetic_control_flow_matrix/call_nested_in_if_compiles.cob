*> vybe-test: cobol/arithmetic_control_flow_matrix/call_nested_in_if_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF F = 1
        CALL "OKM"
    ELSE
        CALL "NOM"
    END-IF.
    STOP RUN.

