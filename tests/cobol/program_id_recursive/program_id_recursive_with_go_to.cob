*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_go_to
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF FLAG = 0
        GO TO DONE
    END-IF.
    DISPLAY "NOT DONE".
DONE.
    DISPLAY "DONE".
    STOP RUN.

