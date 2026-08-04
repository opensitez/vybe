*> vybe-test: cobol/perform_out_of_line/perform_with_test_before_compiles
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM WITH TEST BEFORE UNTIL N >= 5
        ADD 1 TO N
    END-PERFORM.
    STOP RUN.

