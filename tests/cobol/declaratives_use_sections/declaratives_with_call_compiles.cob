*> vybe-test: cobol/declaratives_use_sections/declaratives_with_call_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_declaratives_use_sections.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL.
SELECT WS-FILE ASSIGN TO "ws-file.dat".

DATA DIVISION.
FILE SECTION.
FD WS-FILE.
01 WS-FILE-REC PIC X(80).
PROCEDURE DIVISION.
DECLARATIVES.
D2 SECTION.
    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.
END DECLARATIVES.
M2 SECTION.
    CALL "WORK".
    STOP RUN.

