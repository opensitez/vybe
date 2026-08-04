*> vybe-test: cobol/occurs_indexed_by/occurs_negative_direction_set_and_read
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(2) OCCURS 5 TIMES INDEXED BY IX.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 55 TO E(3).
    SET IX TO 5.
    SET IX DOWN BY 2.
    DISPLAY E(IX).
    MOVE SPACES TO WS-VYBE-L
    STRING E(IX) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "55"
        DISPLAY "FAIL: want [55] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

