*> vybe-test: cobol/date_time_expanded/datetime_evaluate_branch_compiles
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
        WHEN OTHER DISPLAY "N"
    END-EVALUATE.
    STOP RUN.

