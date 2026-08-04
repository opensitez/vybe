*> vybe-test: cobol/category_file_control/test_file_declaratives_error
*> origin: languages/cobol/tests/cobol/test_category_file_control.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. FILE-DECL.
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
       ERR-PROC SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON TEST-FILE.
       ERR-PARA.
           DISPLAY "ERROR CAUGHT " WS-STAT.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "ERROR CAUGHT " DELIMITED SIZE WS-STAT DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "DECLARATIVES PARSED"
                DISPLAY "FAIL at 1 want [DECLARATIVES PARSED] got [" WS-VYBE-L "]"
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
           DISPLAY "DECLARATIVES PARSED".
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "DECLARATIVES PARSED" DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "DECLARATIVES PARSED"
                DISPLAY "FAIL at 1 want [DECLARATIVES PARSED] got [" WS-VYBE-L "]"
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

