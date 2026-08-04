*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_display
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 MSG PIC X(10) VALUE "HELLO".
PROCEDURE DIVISION.
    DISPLAY MSG.
    STOP RUN.

