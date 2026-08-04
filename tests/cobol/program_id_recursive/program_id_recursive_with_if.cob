*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_if
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 5.
PROCEDURE DIVISION.
    IF X > 3
        DISPLAY "BIG"
    ELSE
        DISPLAY "SMALL"
    END-IF.
    STOP RUN.

