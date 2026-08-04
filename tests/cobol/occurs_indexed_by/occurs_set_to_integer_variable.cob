*> vybe-test: cobol/occurs_indexed_by/occurs_set_to_integer_variable
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9 OCCURS 5 TIMES INDEXED BY IX.
01 N PIC 9 VALUE 4.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SET IX TO N.
    MOVE 9 TO E(IX).
    DISPLAY E(4).
    MOVE SPACES TO WS-VYBE-L
    STRING E(4) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "9"
        DISPLAY "FAIL: want [9] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

