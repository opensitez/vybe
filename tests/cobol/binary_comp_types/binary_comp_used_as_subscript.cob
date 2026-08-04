*> vybe-test: cobol/binary_comp_types/binary_comp_used_as_subscript
*> origin: languages/cobol/tests/cobol/test_binary_comp_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 IDX PIC 9(4) COMP VALUE 2.
01 T.
   05 E PIC X OCCURS 5 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "B" TO E(IDX).
    DISPLAY E(IDX).
    MOVE SPACES TO WS-VYBE-L
    STRING E(IDX) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

