*> vybe-test: cobol/numeric_picture_editing/pic_a_alphabetic
*> origin: languages/cobol/tests/cobol/test_numeric_picture_editing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC A(5) VALUE "HELLO".
PROCEDURE DIVISION.
    DISPLAY S.
    STOP RUN.

