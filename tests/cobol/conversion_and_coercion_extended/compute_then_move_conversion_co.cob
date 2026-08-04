*> vybe-test: cobol/conversion_and_coercion_extended/compute_then_move_conversion_compiles
*> origin: languages/cobol/tests/cobol/test_conversion_and_coercion_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 3.
01 WS-B PIC 9(3) VALUE 4.
01 WS-C PIC 9(5).
PROCEDURE DIVISION.
    COMPUTE WS-C = WS-A * WS-B.
    STOP RUN.

