*> vybe-test: cobol/program_id_recursive/program_id_recursive_stop_run_in_branch
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ERR PIC 9 VALUE 0.
PROCEDURE DIVISION.
    IF ERR > 0
        STOP RUN
    END-IF.
    DISPLAY "OK".
    STOP RUN.

