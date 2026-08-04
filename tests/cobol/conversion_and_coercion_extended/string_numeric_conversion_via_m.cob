*> vybe-test: cobol/conversion_and_coercion_extended/string_numeric_conversion_via_move_compiles
*> origin: languages/cobol/tests/cobol/test_conversion_and_coercion_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(3) VALUE "123".
01 WS-NUM PIC 9(3).
PROCEDURE DIVISION.
    MOVE WS-STR TO WS-NUM.
    STOP RUN.

