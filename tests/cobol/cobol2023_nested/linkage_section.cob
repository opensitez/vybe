*> vybe-test: cobol/cobol2023_nested/linkage_section
*> origin: languages/cobol/tests/cobol/test_cobol2023_nested.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-MAIN PIC X(20) VALUE "Main".
LINKAGE SECTION.
01 LS-PARAM PIC X(20).
PROCEDURE DIVISION.
    DISPLAY WS-MAIN.
    STOP RUN.

