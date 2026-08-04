*> vybe-test: cobol/category_data_editing_advanced/compile_edit_redefines_with_edit_pictures
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01  RAW PIC 9(4) VALUE 1234.
01 OUT REDEFINES RAW PIC ZZ,ZZ.
PROCEDURE DIVISION.
    DISPLAY OUT.
    STOP RUN.

