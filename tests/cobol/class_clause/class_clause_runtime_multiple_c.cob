*> vybe-test: cobol/class_clause/class_clause_runtime_multiple_classes
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS14.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS DIGITS IS "0" THRU "9".
    CLASS LETTERS IS "A" THRU "Z".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
01 CH PIC X VALUE "Q".
PROCEDURE DIVISION.
    IF CH IS LETTERS
        DISPLAY "LETTER"
    END-IF
    IF CH IS NOT DIGITS
        DISPLAY "NOTDIG"
    END-IF
    STOP RUN.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "LETTER" DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "LETTER"
                DISPLAY "FAIL at 1 want [LETTER] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "NOTDIG"
                DISPLAY "FAIL at 2 want [NOTDIG] got [" WS-VYBE-L "]"
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

