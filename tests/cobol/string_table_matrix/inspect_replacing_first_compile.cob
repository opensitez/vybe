*> vybe-test: cobol/string_table_matrix/inspect_replacing_first_compiles
*> origin: languages/cobol/tests/cobol/test_string_table_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(6) VALUE "AAAAAA".
PROCEDURE DIVISION.
    INSPECT TXT REPLACING FIRST "A" BY "B".
    STOP RUN.

