*> vybe-test: cobol/advanced_control_flow/evaluate_with_elsif_path_compiles
*> origin: languages/cobol/tests/cobol/test_advanced_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-A PIC 9(1) VALUE 8.
PROCEDURE DIVISION.
    EVALUATE WS-A
        WHEN 1 THRU 3
            DISPLAY "LOW"
        WHEN 4 THRU 7
            DISPLAY "MID"
        WHEN 8 THRU 9
            DISPLAY "HIGH"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "LOW" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HIGH"
        DISPLAY "FAIL: want [HIGH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

