*> vybe-test: cobol/category_perform_varying/test_perf_var_3_levels_test_before
*> origin: languages/cobol/tests/cobol/test_category_perform_varying.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 I PIC 9 VALUE 0. 01 J PIC 9 VALUE 0. 01 K PIC 9 VALUE 0. 01 T PIC 99 VALUE 0. PROCEDURE DIVISION. PERFORM VARYING I FROM 1 BY 1 UNTIL I > 2 AFTER J FROM 1 BY 1 UNTIL J > 2 AFTER K FROM 1 BY 1 UNTIL K > 2 WITH TEST BEFORE ADD 1 TO T END-PERFORM. DISPLAY T.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING T DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "08"
                DISPLAY "FAIL at 1 want [08] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. STOP RUN.
    IF WS-VYBE-I NOT = 1
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 1"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

