*> vybe-test: cobol/collections_tables_expanded/table_indexed_element_display_runtime
*> origin: languages/cobol/tests/cobol/test_collections_tables_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ITEM PIC X(5) OCCURS 2 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "HELLO" TO WS-ITEM(1).
    DISPLAY WS-ITEM(1).
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ITEM(1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO"
        DISPLAY "FAIL: want [HELLO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

