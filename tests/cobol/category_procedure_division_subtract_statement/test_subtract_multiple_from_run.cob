*> vybe-test: cobol/category_procedure_division_subtract_statement/test_subtract_multiple_from_runtime
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_subtract_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 V1 PIC 9 VALUE 10. 01 V2 PIC 9 VALUE 1. 01 V3 PIC 9 VALUE 2. 01 V4 PIC 9 VALUE 3. PROCEDURE DIVISION. SUBTRACT V2 V3 FROM V1 V4 DISPLAY V1 DISPLAY V4 STOP RUN.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING V1 DELIMITED SIZE DISPLAY DELIMITED SIZE V4 DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "5"
                DISPLAY "FAIL at 1 want [5] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "3"
                DISPLAY "FAIL at 2 want [3] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 2 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    IF WS-VYBE-I NOT = 2
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 2"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

