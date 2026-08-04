*> vybe-test: cobol/exceptions_error_paths/fallback_path_with_evaluate_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ST PIC 9 VALUE 3.
PROCEDURE DIVISION.
    EVALUATE ST
        WHEN 1 DISPLAY "OK"
        WHEN 2 DISPLAY "WARN"
        WHEN OTHER DISPLAY "ERR"
    END-EVALUATE.
    STOP RUN.

