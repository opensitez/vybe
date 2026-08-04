*> vybe-test: cobol/cobol/evaluate_when
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. EVAL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GRADE PIC X(1) VALUE "B".
PROCEDURE DIVISION.
    EVALUATE WS-GRADE
        WHEN "A"
            DISPLAY "Excellent"
        WHEN "B"
            DISPLAY "Good"
        WHEN "C"
            DISPLAY "Average"
        WHEN OTHER
            DISPLAY "Unknown"
    END-EVALUATE.
    STOP RUN.

