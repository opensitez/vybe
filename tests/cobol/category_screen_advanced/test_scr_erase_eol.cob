*> vybe-test: cobol/category_screen_advanced/test_scr_erase_eol
*> vybe-test-mode: compile
*> A SCREEN SECTION program runs under curses: every DISPLAY — including
*> the generated checker's own — goes to the terminal, not to stdout. cobc
*> compiles this and then fails the SAME embedded assertion (measured
*> 2026-08-29: 33 of 33 in this suite exit non-zero under cobc), so the
*> line-counting checker is not a property this source can have under any
*> COBOL. Compiling is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_category_screen_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. SCREEN SECTION. 01 S1. 05 LINE 1 COL 1 VALUE 'A' ERASE EOL. PROCEDURE DIVISION. DISPLAY S1.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING S1 DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "A"
                DISPLAY "FAIL at 1 want [A] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. DISPLAY 'OK'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'OK' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "OK"
                DISPLAY "FAIL at 1 want [OK] got [" WS-VYBE-L "]"
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

