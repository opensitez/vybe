*> vybe-test: cobol/string_functions_intrinsic/intrinsic_when_compiled_compiles_and_displays
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 COMPILED-AT PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION WHEN-COMPILED TO COMPILED-AT.
    DISPLAY COMPILED-AT.
    STOP RUN.

