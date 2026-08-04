*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_perform_varying
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(3) VALUE 0.
01 S PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 100
        ADD I TO S
    END-PERFORM.
    STOP RUN.

