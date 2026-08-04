*> vybe-test: cobol/redefines_extended/redefines_with_group_item_is_accepted
*> origin: languages/cobol/tests/cobol/test_redefines_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-FIELD1 PIC X(2) VALUE "AA".
   05 WS-FIELD2 PIC X(2) VALUE "BB".
01 WS-ALT REDEFINES WS-REC.
   05 WS-VAL PIC X(4).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    DISPLAY WS-VAL.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-VAL DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AABB"
        DISPLAY "FAIL: want [AABB] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

