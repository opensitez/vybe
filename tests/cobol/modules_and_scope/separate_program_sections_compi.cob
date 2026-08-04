*> vybe-test: cobol/modules_and_scope/separate_program_sections_compile
*> origin: languages/cobol/tests/cobol/test_modules_and_scope.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(1) VALUE 1.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

