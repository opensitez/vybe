*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_evaluate
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE N
        WHEN 1 DISPLAY "ONE"
        WHEN 2 DISPLAY "TWO"
        WHEN OTHER DISPLAY "OTHER"
    END-EVALUATE.
    STOP RUN.

