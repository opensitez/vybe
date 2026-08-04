*> vybe-test: cobol/category_data_division_occurs/test_occurs_reference_modification
*> origin: languages/cobol/tests/cobol/test_category_data_division_occurs.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 E OCCURS 3 TIMES PIC X(3). PROCEDURE DIVISION. MOVE 'ABC' TO E(1) MOVE 'DEF' TO E(2) MOVE 'GHI' TO E(3) DISPLAY E(2)(2:2) STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING E(2)(2:2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "EF"
        DISPLAY "FAIL: want [EF] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

