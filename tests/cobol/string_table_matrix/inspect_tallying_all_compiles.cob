*> vybe-test: cobol/string_table_matrix/inspect_tallying_all_compiles
*> origin: languages/cobol/tests/cobol/test_string_table_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TXT PIC X(8) VALUE "ABABXABA".
01 C PIC 9(2) VALUE 0.
PROCEDURE DIVISION.
    INSPECT TXT TALLYING C FOR ALL "A".
    STOP RUN.

