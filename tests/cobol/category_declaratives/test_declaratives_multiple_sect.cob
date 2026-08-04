*> vybe-test: cobol/category_declaratives/test_declaratives_multiple_sections_runtime
*> origin: languages/cobol/tests/cobol/test_category_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DECL-MULTI.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT TEST-FILE ASSIGN TO "err.dat"
           FILE STATUS IS WS-STAT.
       DATA DIVISION.
       FILE SECTION.
       FD TEST-FILE.
       01 REC PIC X.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
       01 WS-STAT PIC XX.
       PROCEDURE DIVISION.
       DECLARATIVES.
       FILE-ERR SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON TEST-FILE.
       FILE-PARA.
           DISPLAY "FILE ERROR".
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "FILE ERROR" DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "MAIN"
                DISPLAY "FAIL at 1 want [MAIN] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
       GLOBAL-ERR SECTION.
           USE GLOBAL AFTER STANDARD ERROR PROCEDURE ON TEST-FILE.
       GLOBAL-PARA.
           DISPLAY "GLOBAL ERROR".
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "GLOBAL ERROR" DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "MAIN"
                DISPLAY "FAIL at 1 want [MAIN] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
       END DECLARATIVES.
       MAIN SECTION.
           DISPLAY "MAIN".
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "MAIN" DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "MAIN"
                DISPLAY "FAIL at 1 want [MAIN] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
           STOP RUN.
       
    IF WS-VYBE-I NOT = 1
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 1"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

