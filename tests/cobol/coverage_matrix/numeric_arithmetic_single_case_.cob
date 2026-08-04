*> vybe-test: cobol/coverage_matrix/numeric_arithmetic_single_case_runtime
*> origin: languages/cobol/tests/cobol/test_coverage_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SAMPLE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
01 A PIC 9(2) VALUE 2.
01 B PIC 9(2) VALUE 3.
01 R PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    ADD A TO B
    DISPLAY B
    SUBTRACT A FROM B
    DISPLAY B
    STOP RUN.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING B DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "5"
                DISPLAY "FAIL at 1 want [5] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "3"
                DISPLAY "FAIL at 2 want [3] got [" WS-VYBE-L "]"
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

