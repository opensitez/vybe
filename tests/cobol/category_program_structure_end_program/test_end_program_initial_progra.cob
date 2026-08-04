*> vybe-test: cobol/category_program_structure_end_program/test_end_program_initial_program_reused
*> origin: languages/cobol/tests/cobol/test_category_program_structure_end_program.rs
IDENTIFICATION DIVISION. PROGRAM-ID. ROOT. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 C PIC 9 VALUE 0. PROCEDURE DIVISION. CALL 'WORKER'. CALL 'WORKER'. STOP RUN. IDENTIFICATION DIVISION. PROGRAM-ID. WORKER IS INITIAL PROGRAM. DATA DIVISION. WORKING-STORAGE SECTION. 01 V PIC 9 VALUE 0. PROCEDURE DIVISION. ADD 1 TO V. DISPLAY V.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING V DELIMITED SIZE INTO WS-VYBE-L
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
        WHEN OTHER
            DISPLAY "FAIL: more than 2 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. EXIT PROGRAM. END PROGRAM WORKER.
    IF WS-VYBE-I NOT = 2
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 2"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

