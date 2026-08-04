*> vybe-test: cobol/value_all_forms/value_all_in_group_item
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 GRP.
   05 P1 PIC X(3) VALUE ALL "A".
   05 P2 PIC X(3) VALUE ALL "B".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY GRP.
    MOVE SPACES TO WS-VYBE-L
    STRING GRP DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AAABBB"
        DISPLAY "FAIL: want [AAABBB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

