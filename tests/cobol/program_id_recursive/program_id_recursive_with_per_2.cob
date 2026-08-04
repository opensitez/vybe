*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_perform_n_times
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 C PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM 10 TIMES
        ADD 1 TO C
    END-PERFORM.
    STOP RUN.

