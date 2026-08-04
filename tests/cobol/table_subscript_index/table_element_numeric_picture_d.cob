*> vybe-test: cobol/table_subscript_index/table_element_numeric_picture_displays_zeroes
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(4) OCCURS 3 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY E(2).
    MOVE SPACES TO WS-VYBE-L
    STRING E(2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0000"
        DISPLAY "FAIL: want [0000] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

