*> vybe-test: cobol/cobol2023_intrinsics/formatted_time
*> origin: languages/cobol/tests/cobol/test_cobol2023_intrinsics.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TIME PIC X(8).
PROCEDURE DIVISION.
    MOVE FUNCTION FORMATTED-TIME("HH:MM:SS" 123000) TO WS-TIME.
    DISPLAY WS-TIME.
    STOP RUN.

