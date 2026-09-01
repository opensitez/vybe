*> vybe-test: cobol/arrays_tables_indexing/table_compute_element_compiles
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC 9(3) OCCURS 2 TIMES.
01 X PIC 9(3) VALUE 4.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE T(1) = X + 2.
    DISPLAY T(1).
    MOVE SPACES TO WS-VYBE-L
    STRING T(1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "006"
        DISPLAY "FAIL: want [006] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

