*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_nested_if
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    IF A > 0
        IF B > 0
            DISPLAY "BOTH"
        END-IF
    END-IF.
    STOP RUN.

