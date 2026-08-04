*> vybe-test: cobol/data_division_extended/data_division_redefines_clause_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BUF PIC X(4) VALUE "ABCD".
01 WS-NUM REDEFINES WS-BUF PIC 9(4).
PROCEDURE DIVISION.
    DISPLAY WS-NUM.
    STOP RUN.

