*> vybe-test: cobol/redefines_extended/redefines_preserves_storage_when_moving_data
*> origin: languages/cobol/tests/cobol/test_redefines_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BUFFER PIC X(4) VALUE "ABCD".
01 WS-HEX REDEFINES WS-BUFFER PIC X(4).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    DISPLAY WS-HEX.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-HEX DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCD"
        DISPLAY "FAIL: want [ABCD] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

