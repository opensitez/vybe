*> vybe-test: cobol/table_subscript_index/table_reference_modification_in_element
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X(6) OCCURS 3 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "ABCDEF" TO E(1).
    DISPLAY E(1)(2:3).
    MOVE SPACES TO WS-VYBE-L
    STRING E(1)(2:3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BCD"
        DISPLAY "FAIL: want [BCD] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

