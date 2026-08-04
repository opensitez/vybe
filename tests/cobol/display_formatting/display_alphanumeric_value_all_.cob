*> vybe-test: cobol/display_formatting/display_alphanumeric_value_all_literal
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(6) VALUE ALL "XY".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XYXYXY"
        DISPLAY "FAIL: want [XYXYXY] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

