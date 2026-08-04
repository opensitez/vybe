*> vybe-test: cobol/date_time_expanded/function_current_date_move_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 CD PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE TO CD.
    STOP RUN.

