*> vybe-test: cobol/inspect_advanced/test_inspect_before_initial
*> origin: languages/cobol/tests/cobol/test_inspect_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(10) VALUE "ABC,DEF,GHI".
01 WS-CNT PIC 9(3) VALUE 0.
PROCEDURE DIVISION.

    INSPECT WS-STR TALLYING WS-CNT FOR ALL "D" BEFORE INITIAL ",".
    STOP RUN.

