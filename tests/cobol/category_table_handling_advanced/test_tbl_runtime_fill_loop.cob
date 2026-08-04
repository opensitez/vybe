*> vybe-test: cobol/category_table_handling_advanced/test_tbl_runtime_fill_loop
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 TBL. 05 EL OCCURS 3 TIMES PIC 9(2). 01 I PIC 99 VALUE 1. PROCEDURE DIVISION. MOVE 10 TO EL(1) MOVE 20 TO EL(2) MOVE 30 TO EL(3) PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3 COMPUTE EL(I) = EL(I) + 1 END-PERFORM DISPLAY EL(1) DISPLAY EL(2) DISPLAY EL(3).
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING EL(1) DELIMITED SIZE DISPLAY DELIMITED SIZE EL(2) DELIMITED SIZE DISPLAY DELIMITED SIZE EL(3) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "11"
                DISPLAY "FAIL at 1 want [11] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "21"
                DISPLAY "FAIL at 2 want [21] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "31"
                DISPLAY "FAIL at 3 want [31] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 3 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. STOP RUN.
    IF WS-VYBE-I NOT = 3
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 3"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

