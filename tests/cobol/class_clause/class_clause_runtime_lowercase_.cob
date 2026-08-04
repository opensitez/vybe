*> vybe-test: cobol/class_clause/class_clause_runtime_lowercase_condition
*> origin: languages/cobol/tests/cobol/test_class_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CLS12.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS DIGIT-CLASS IS "0" THRU "9".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
01 WS-NUM PIC X VALUE "7".
01 WS-CHAR PIC X VALUE "A".
PROCEDURE DIVISION.
    IF WS-NUM IS DIGIT-CLASS
        DISPLAY "NUM"
    END-IF
    IF WS-CHAR IS NOT DIGIT-CLASS
        DISPLAY "NOTNUM"
    END-IF
    STOP RUN.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "NUM" DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "NUM"
                DISPLAY "FAIL at 1 want [NUM] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "NOTNUM"
                DISPLAY "FAIL at 2 want [NOTNUM] got [" WS-VYBE-L "]"
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

