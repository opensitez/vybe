*> vybe-test: cobol/category_data_editing_advanced/compile_edit_picture_v_significant_zeros
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC ZZ,ZZ9.
01 Y PIC 9999 VALUE 100.
PROCEDURE DIVISION.
    MOVE Y TO X.
    STOP RUN.

