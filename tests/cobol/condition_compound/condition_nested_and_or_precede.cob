*> vybe-test: cobol/condition_compound/condition_nested_and_or_precedence
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 0.
01 C PIC 9 VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A = 1 AND B = 1 OR C = 1
        DISPLAY "TRUE"
    ELSE
        DISPLAY "FALSE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "TRUE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "TRUE"
        DISPLAY "FAIL: want [TRUE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

