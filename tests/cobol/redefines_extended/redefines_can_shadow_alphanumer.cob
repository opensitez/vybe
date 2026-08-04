*> vybe-test: cobol/redefines_extended/redefines_can_shadow_alphanumeric_value_with_numeric_view
*> origin: languages/cobol/tests/cobol/test_redefines_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BUFFER PIC X(4) VALUE "1234".
01 WS-NUMBER REDEFINES WS-BUFFER PIC 9(4).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE 9999 TO WS-NUMBER.
    DISPLAY WS-BUFFER.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-BUFFER DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "9999"
        DISPLAY "FAIL: want [9999] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

