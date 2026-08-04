*> vybe-test: cobol/occurs_indexed_by/occurs_subscript_boundary_first
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X OCCURS 3 TIMES INDEXED BY IX.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "A" TO E(1).
    DISPLAY E(1).
    MOVE SPACES TO WS-VYBE-L
    STRING E(1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A"
        DISPLAY "FAIL: want [A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

