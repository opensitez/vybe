*> vybe-test: cobol/date_time_expanded/day_of_week_evaluate_full_branch_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 W PIC X(1).
PROCEDURE DIVISION.
    ACCEPT W FROM DAY-OF-WEEK.
    EVALUATE W
        WHEN "1" DISPLAY "MON"
        WHEN "2" DISPLAY "TUE"
        WHEN OTHER DISPLAY "X"
    END-EVALUATE.
    STOP RUN.

