*> vybe-test: cobol/arithmetic_operations_matrix/div_case_runtime_quotient_remainder
*> origin: languages/cobol/tests/cobol/test_arithmetic_operations_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 99 VALUE 20.
01 B PIC 9 VALUE 3.
01 Q PIC 99.
01 M PIC 9.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DIVIDE B INTO A GIVING Q REMAINDER M.
    DISPLAY Q.
    MOVE SPACES TO WS-VYBE-L
    STRING Q DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "06"
        DISPLAY "FAIL: want [06] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY M.
    MOVE SPACES TO WS-VYBE-L
    STRING M DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

