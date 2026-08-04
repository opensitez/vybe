*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_loop
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL I >= 10
        ADD 1 TO I
    END-PERFORM.
    STOP RUN.

