*> vybe-test: cobol/programs/grade_calculator
*> origin: languages/cobol/tests/cobol/test_programs.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. GRADES.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SCORE PIC 9(3) VALUE 0.
01 WS-TOTAL PIC 9(5) VALUE 0.
01 WS-COUNT PIC 9(3) VALUE 5.
01 WS-AVG   PIC 9(5)V99 VALUE 0.
01 WS-I     PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    MOVE 0 TO WS-TOTAL.
    ADD 85 TO WS-TOTAL.
    ADD 92 TO WS-TOTAL.
    ADD 78 TO WS-TOTAL.
    ADD 95 TO WS-TOTAL.
    ADD 88 TO WS-TOTAL.
    COMPUTE WS-AVG = WS-TOTAL / WS-COUNT.
    DISPLAY "Average: " WS-AVG.
    EVALUATE TRUE
        WHEN WS-AVG >= 90
            DISPLAY "Grade: A"
        WHEN WS-AVG >= 80
            DISPLAY "Grade: B"
        WHEN WS-AVG >= 70
            DISPLAY "Grade: C"
        WHEN WS-AVG >= 60
            DISPLAY "Grade: D"
        WHEN OTHER
            DISPLAY "Grade: F"
    END-EVALUATE.
    STOP RUN.

