*> vybe-test: cobol/evaluate_advanced/test_evaluate_function_subject
*> origin: languages/cobol/tests/cobol/test_evaluate_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    EVALUATE FUNCTION MOD(WS-X 3)
        WHEN 0
            DISPLAY "DIV3"
        WHEN 2
            DISPLAY "REM2"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "DIV3" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "DIV3"
        DISPLAY "FAIL: want [DIV3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

