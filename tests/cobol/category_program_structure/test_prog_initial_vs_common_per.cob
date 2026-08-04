*> vybe-test: cobol/category_program_structure/test_prog_initial_vs_common_persist
*> origin: languages/cobol/tests/cobol/test_category_program_structure.rs
IDENTIFICATION DIVISION. PROGRAM-ID. ROOT. PROCEDURE DIVISION. CALL 'COUNTER'. CALL 'COUNTER'. CALL 'COMMON'. CALL 'COMMON'. STOP RUN. IDENTIFICATION DIVISION. PROGRAM-ID. COUNTER IS INITIAL PROGRAM. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 C PIC 9 VALUE 0. PROCEDURE DIVISION. ADD 1 TO C DISPLAY C.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING C DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "1"
                DISPLAY "FAIL at 1 want [1] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "1"
                DISPLAY "FAIL at 2 want [1] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "1"
                DISPLAY "FAIL at 3 want [1] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "2"
                DISPLAY "FAIL at 4 want [2] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 4 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. EXIT PROGRAM. END PROGRAM COUNTER. IDENTIFICATION DIVISION. PROGRAM-ID. COMMON IS COMMON PROGRAM. DATA DIVISION. WORKING-STORAGE SECTION. 01 C2 PIC 9 VALUE 0. PROCEDURE DIVISION. ADD 1 TO C2 DISPLAY C2.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING C2 DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "1"
                DISPLAY "FAIL at 1 want [1] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "1"
                DISPLAY "FAIL at 2 want [1] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "1"
                DISPLAY "FAIL at 3 want [1] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "2"
                DISPLAY "FAIL at 4 want [2] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 4 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. EXIT PROGRAM. END PROGRAM COMMON.
    IF WS-VYBE-I NOT = 4
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 4"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

