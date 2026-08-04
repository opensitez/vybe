*> vybe-test: cobol/data_division_expanded/sync_clause_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-A PIC X(3).
   05 WS-B PIC X(3) SYNC.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "ABC" TO WS-A.
    MOVE "XYZ" TO WS-B.
    DISPLAY WS-B.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XYZ"
        DISPLAY "FAIL: want [XYZ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

