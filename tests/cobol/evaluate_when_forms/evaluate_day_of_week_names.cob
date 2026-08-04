*> vybe-test: cobol/evaluate_when_forms/evaluate_day_of_week_names
*> origin: languages/cobol/tests/cobol/test_evaluate_when_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DAY PIC 9 VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE DAY
        WHEN 1
            DISPLAY "MON"
        WHEN 2
            DISPLAY "TUE"
        WHEN 3
            DISPLAY "WED"
        WHEN 4
            DISPLAY "THU"
        WHEN 5
            DISPLAY "FRI"
        WHEN OTHER
            DISPLAY "WKD"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "MON" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MON"
        DISPLAY "FAIL: want [MON] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

