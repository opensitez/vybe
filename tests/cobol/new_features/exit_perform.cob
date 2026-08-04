*> vybe-test: cobol/new_features/exit_perform
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(10) VALUE 0.
01 A PIC 9(10) VALUE 10.
01 B PIC 9(10) VALUE 20.
PROCEDURE DIVISION.
    PERFORM 10 TIMES
        EXIT PERFORM
    END-PERFORM.
    STOP RUN.

