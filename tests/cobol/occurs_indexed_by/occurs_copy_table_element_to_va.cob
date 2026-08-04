*> vybe-test: cobol/occurs_indexed_by/occurs_copy_table_element_to_variable
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X(5) OCCURS 3 TIMES INDEXED BY IX.
01 COPY PIC X(5) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "HELLO" TO E(2).
    SET IX TO 2.
    MOVE E(IX) TO COPY.
    DISPLAY COPY.
    MOVE SPACES TO WS-VYBE-L
    STRING COPY DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO"
        DISPLAY "FAIL: want [HELLO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

