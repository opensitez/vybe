*> vybe-test: cobol/category_data_editing_advanced/compile_edit_alphanumeric_editing_clauses
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(10).
01 Y PIC 9(4) VALUE 100.
PROCEDURE DIVISION.
    MOVE Y TO X.
    STOP RUN.

