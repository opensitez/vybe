*> vybe-test: cobol/category_declaratives/test_declaratives_for_debugging_on_section_runtime
*> origin: languages/cobol/tests/cobol/test_category_declaratives.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DECL-DBG.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       SOURCE-COMPUTER. COMPUTER WITH DEBUGGING MODE.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
       01 WS-VAL PIC 9 VALUE 1.
       PROCEDURE DIVISION.
       DECLARATIVES.
       DBG-SEC SECTION.
           USE FOR DEBUGGING ON MAIN-SECTION.
       DBG-PARA.
           DISPLAY "DBG".
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "DBG" DELIMITED SIZE INTO WS-VYBE-L
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
       MAIN-SECTION.
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

